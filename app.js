/**
 * 學生評分系統 - 主應用程式
 */

// 全域狀態
let currentUser = null;
let teacherInfo = null;
let tokenClient = null;
let gapiInited = false;
let gisInited = false;

// 資料快取
let teachersData = [];
let studentsData = [];
let relationsData = [];
let scoresData = [];
let exportsData = []; // 資料快取：評分匯出表
let dataLoaded = false; // whether loadAllData has successfully loaded data

/**
 * 判斷是否為 Sheets 權限錯誤（寫入 / 編輯 權限不足）
 */
function isSheetsPermissionError(err) {
    if (!err) {
        return false;
    }

    // gapi 回傳的錯誤可能包含 status 或 result.error.code / message
    const code = err.status || (err.result && err.result.error && err.result.error.code) || (err.statusCode);
    const msg = (err.result && err.result.error && err.result.error.message) || err.message || '';
    if (code === 403 || /permission|insufficient|does not have permission|not authorized/i.test(msg)) return true;
    return false;
}

// HTML 屬性值跳脫（防止 data-* 屬性中出現 " 破壞 HTML 結構）
function escapeAttr(str) {
    return String(str == null ? '' : str)
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;');
}

// DOM 元素
const elements = {};

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    initElements();
    initEventListeners();
    loadGoogleAPI();
});

// Session persistence keys
const SESSION_KEY = 'grading_app_session_v1';

/**
 * 初始化 DOM 元素引用
 */
function initElements() {
    elements.loginSection = document.getElementById('login-section');
    elements.mainSection = document.getElementById('main-section');
    elements.loginBtn = document.getElementById('login-btn');
    elements.logoutBtn = document.getElementById('logout-btn');
    elements.loginError = document.getElementById('login-error');
    elements.userInfo = document.getElementById('user-info');
    
    // 標籤頁
    elements.tabBtns = document.querySelectorAll('.tab-btn');
    elements.scoringTab = document.getElementById('scoring-tab');
    elements.managementTab = document.getElementById('management-tab');
    
    // 學生列表
    elements.studentsLoading = document.getElementById('students-loading');
    elements.studentsEmpty = document.getElementById('students-empty');
    elements.studentsList = document.getElementById('students-list');
    
    // 關係管理
    elements.newTeacher = document.getElementById('new-teacher');
    elements.newStudent = document.getElementById('new-student');
    elements.addRelationBtn = document.getElementById('add-relation-btn');
    elements.newTeacherName = document.getElementById('new-teacher-name');
    elements.newTeacherEmail = document.getElementById('new-teacher-email');
    elements.addTeacherBtn = document.getElementById('add-teacher-btn');
    elements.newStudentName = document.getElementById('new-student-name');
    elements.newStudentNumber = document.getElementById('new-student-number');
    elements.addStudentBtn = document.getElementById('add-student-btn');
    elements.relationsLoading = document.getElementById('relations-loading');
    elements.relationsTable = document.getElementById('relations-table');
    elements.relationsTbody = document.getElementById('relations-tbody');
    
    // Modal
    elements.scoreModal = document.getElementById('score-modal');
    elements.modalStudentName = document.getElementById('modal-student-name');
    elements.scoreInput = document.getElementById('score-input');
    elements.commentInput = document.getElementById('comment-input');
    elements.cancelScoreBtn = document.getElementById('cancel-score-btn');
    elements.submitScoreBtn = document.getElementById('submit-score-btn');
    
    // Toast
    elements.toast = document.getElementById('toast');
}

/**
 * 初始化事件監聽器
 */
function initEventListeners() {
    // 登入/登出
    elements.loginBtn.addEventListener('click', handleLogin);
    elements.logoutBtn.addEventListener('click', handleLogout);
    
    // 標籤頁切換
    elements.tabBtns.forEach(btn => {
        btn.addEventListener('click', () => switchTab(btn.dataset.tab));
    });
    
    // 關係管理
    elements.addRelationBtn.addEventListener('click', handleAddRelation);
    elements.addTeacherBtn.addEventListener('click', handleAddTeacher);
    elements.addStudentBtn.addEventListener('click', handleAddStudent);
    
    // Modal
    elements.cancelScoreBtn.addEventListener('click', closeScoreModal);
    elements.submitScoreBtn.addEventListener('click', handleSubmitScore);
    elements.scoreModal.querySelector('.modal-backdrop').addEventListener('click', closeScoreModal);
    elements.scoreModal.querySelector('.modal-close').addEventListener('click', closeScoreModal);
}

/**
 * 載入 Google API
 */
function loadGoogleAPI() {
    // 載入 GAPI
    const gapiScript = document.createElement('script');
    gapiScript.src = 'https://apis.google.com/js/api.js';
    gapiScript.onload = gapiLoaded;
    document.head.appendChild(gapiScript);
    
    // 載入 GIS (Google Identity Services)
    const gisScript = document.createElement('script');
    gisScript.src = 'https://accounts.google.com/gsi/client';
    gisScript.onload = gisLoaded;
    document.head.appendChild(gisScript);
}

/**
 * GAPI 載入完成
 */
function gapiLoaded() {
    gapi.load('client', async () => {
        await gapi.client.init({
            discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
        });
        gapiInited = true;
        maybeEnableLogin();
    });
}

/**
 * GIS 載入完成
 */
function gisLoaded() {
    tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CONFIG.GOOGLE_CLIENT_ID,
        scope: CONFIG.SCOPES,
        callback: handleAuthCallback,
    });
    gisInited = true;
    maybeEnableLogin();
}

/**
 * 檢查是否可以啟用登入按鈕
 */
async function maybeEnableLogin() {
    if (gapiInited && gisInited) {
        elements.loginBtn.disabled = false;

        // Attempt to restore session before showing any UI to avoid flash
        // overlay is visible by default; we'll hide it after restore attempt
        const restored = await restoreSessionIfAny();
        // Hide app loading overlay
        hideAppLoading();

        if (restored) {
            // restoreSession handled setting currentUser & teacherInfo
            showMainSection();
        } else {
            // No valid saved session -> show login UI
            showLoginSection();
        }
    }
}

/**
 * 處理登入
 */
function handleLogin() {
    if (gapi.client.getToken() === null) {
        tokenClient.requestAccessToken({ prompt: 'consent' });
    } else {
        tokenClient.requestAccessToken({ prompt: '' });
    }
}

/**
 * 處理認證回調
 */
async function handleAuthCallback(response) {
    if (response.error !== undefined) {
        showError('登入失敗：' + response.error);
        return;
    }
    
    // 獲取使用者資訊
    try {
        const userInfoResponse = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
            headers: {
                'Authorization': `Bearer ${gapi.client.getToken().access_token}`
            }
        });
        currentUser = await userInfoResponse.json();

        // 若 userinfo 沒有 email（通常是因為沒有請求 openid/email/profile），提示並嘗試要求額外授權
        if (!currentUser || !currentUser.email) {
            console.warn('userinfo 沒有回傳 email', currentUser);
            showError('登入成功但無法取得 Email。請確認 OAuth scopes 包含 openid、email、profile，並再次同意授權。');
            // 觸發要求使用者再次授權，帶出同意畫面
            try {
                tokenClient.requestAccessToken({ prompt: 'consent' });
            } catch (e) {
                console.warn('無法要求額外授權：', e);
            }
            return;
        }

        // 驗證使用者是否為老師
        await loadTeachersData();
        teacherInfo = teachersData.find(t => t.email.toLowerCase() === currentUser.email.toLowerCase());
        
        if (!teacherInfo) {
            showError(`此帳號 (${currentUser.email}) 不在老師名單中，無法使用系統。`);
            handleLogout();
            return;
        }
        
        // 嘗試載入所有資料（同時能檢查是否有存取權限）
        const ok = await loadAllData();
        if (!ok) {
            // loadAllData 已處理顯示錯誤與登出 (如適用)
            return;
        }

        // 儲存 session（access token 與 userinfo）以在重新整理時保留登入狀態
        try {
            const savedToken = gapi.client.getToken();
            saveSession({ token: savedToken, user: currentUser });
        } catch (e) {
            console.warn('無法儲存 session token:', e);
        }

        // 登入成功，顯示主介面
        showMainSection();
        
    } catch (error) {
        console.error('Error getting user info:', error);
        if (isSheetsPermissionError(error)) {
            showError('您沒有權限使用本系統或其資料，請使用具有權限的帳號或聯絡管理員。');
            // remove saved session and logout to allow trying another account
            clearSavedSession();
            handleLogout();
            return;
        }
        showError('無法獲取使用者資訊');
    }
}

/**
 * 處理登出
 */
function handleLogout() {
    const token = gapi.client.getToken();
    if (token !== null) {
        google.accounts.oauth2.revoke(token.access_token);
        gapi.client.setToken('');
    }
    currentUser = null;
    teacherInfo = null;
    // remove stored session
    clearSavedSession();
    showLoginSection();
}

/**
 * 儲存 session 到 localStorage
 */
function saveSession(sessionObj) {
    try {
        localStorage.setItem(SESSION_KEY, JSON.stringify(sessionObj));
    } catch (e) {
        console.warn('saveSession failed', e);
    }
}

/**
 * 清除 localStorage 的 session
 */
function clearSavedSession() {
    try { localStorage.removeItem(SESSION_KEY); } catch(e) { /* ignore */ }
}

/**
 * 嘗試從 localStorage 還原 session
 */
async function restoreSessionIfAny() {
    try {
        const raw = localStorage.getItem(SESSION_KEY);
        if (!raw) return false;

        const obj = JSON.parse(raw);
        if (!obj || !obj.token || !obj.token.access_token) return false;

        // 將 token 加到 gapi client
        gapi.client.setToken(obj.token);

        // 驗證 token 是否仍然有效：呼叫 userinfo
        const userInfoResponse = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
            headers: { 'Authorization': `Bearer ${obj.token.access_token}` }
        });

        if (!userInfoResponse.ok) {
            // token 不可用，清除
            clearSavedSession();
            return false;
        }

        const userinfo = await userInfoResponse.json();
        if (!userinfo || !userinfo.email) {
            clearSavedSession();
            return false;
        }

        // 恢復狀態並檢查老師清單
        currentUser = userinfo;
        await loadTeachersData();
        teacherInfo = teachersData.find(t => t.email.toLowerCase() === currentUser.email.toLowerCase());
        if (!teacherInfo) {
            // token 有效，但該帳號不在老師名單：清除 session，回報失敗 so caller can show login and message
            clearSavedSession();
            // keep user info to show a helpful message
            showError(`此帳號 (${currentUser.email}) 不在老師名單中，無法使用系統。`);
            return false;
        }

        // 若一切正常，回傳成功，caller 決定要顯示主畫面
        return true;
    } catch (err) {
        console.warn('restoreSessionIfAny failed', err);
        if (isSheetsPermissionError(err)) {
            // show clear permission message and log out so user can try another account
            showError('您沒有權限使用本系統或其資料，請使用具有權限的帳號或聯絡管理員。');
            clearSavedSession();
            handleLogout();
            return false;
        }
        // 若發生任何其他錯誤，清除 session
        clearSavedSession();
        return false;
    }
}

function showAppLoading() {
    const el = document.getElementById('app-loading');
    if (el) el.classList.remove('hidden');
}

function hideAppLoading() {
    const el = document.getElementById('app-loading');
    if (el) el.classList.add('hidden');
}

/**
 * 顯示登入區域
 */
function showLoginSection() {
    hideAppLoading();
    elements.mainSection.classList.add('hidden');
    elements.loginSection.classList.remove('hidden');
}

/**
 * 顯示主要區域
 */
async function showMainSection() {
    hideAppLoading();
    // clear any login error banner when successfully showing main UI
    if (elements.loginError) elements.loginError.classList.add('hidden');
    elements.loginSection.classList.add('hidden');
    elements.mainSection.classList.remove('hidden');
    elements.userInfo.textContent = `${teacherInfo.name} (${currentUser.email})`;
    
    // 載入資料（若還沒載入過）
    if (!dataLoaded) {
        await loadAllData();
    }
    renderStudentsList();
    renderManagementTab();
}

/**
 * 顯示錯誤訊息
 */
function showError(message) {
    elements.loginError.textContent = message;
    elements.loginError.classList.remove('hidden');
}

/**
 * 切換標籤頁
 */
function switchTab(tabName) {
    elements.tabBtns.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabName);
    });
    
    elements.scoringTab.classList.toggle('hidden', tabName !== 'scoring');
    elements.scoringTab.classList.toggle('active', tabName === 'scoring');
    elements.managementTab.classList.toggle('hidden', tabName !== 'management');
    elements.managementTab.classList.toggle('active', tabName === 'management');
}

/**
 * 載入所有資料
 */
async function loadAllData() {
    try {
        await Promise.all([
            loadTeachersData(),
            loadStudentsData(),
            loadRelationsData(),
            loadScoresData(),
            loadExportsData()
        ]);
        dataLoaded = true;
        return true;
    } catch (error) {
        console.error('Error loading data:', error);
        if (isSheetsPermissionError(error)) {
            // Show a clearer message and log user out so they can try with a different Google account
            showError('您沒有權限使用本系統或其資料，請使用具有權限的帳號或聯絡管理員。');
            // clear any saved session and force logout
            clearSavedSession();
            handleLogout();
            return false;
        }
        showToast('載入資料失敗', 'error');
        return false;
    }
}

/**
 * 載入匯出表格（評分匯出）
 * 欄位預期: A: 老師姓名, B: 學生姓名, C: 分數, D: 評語, E: 評分時間
 */
async function loadExportsData() {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: CONFIG.SPREADSHEET_ID,
            range: `${CONFIG.SHEETS.EXPORTS}!A2:E`,
        });

        const values = response.result.values || [];
        exportsData = values.map((row, index) => ({
            rowIndex: index + 2,
            teacherName: row[0] || '',
            studentName: row[1] || '',
            score: row[2] || '',
            comment: row[3] || '',
            timestamp: row[4] || ''
        }));
    } catch (error) {
        console.error('Error loading exports:', error);
        if (isSheetsPermissionError(error)) throw error;
        exportsData = [];
    }
}

/**
 * 根據 scoresData 重新生成評分匯出表的內容 (A:E)
 * 會先清除 A2:E 範圍，再 append 所有資料，確保與 scores 對應
 */
async function syncExportsFromScores() {
    try {
        // build values from scoresData in order
        const values = scoresData.map(s => {
            const teacher = teachersData.find(t => t.email.toLowerCase() === (s.teacherEmail || '').toLowerCase());
            const student = studentsData.find(st => st.id === s.studentId);
            const teacherName = teacher ? teacher.name : (s.teacherEmail || '');
            const studentName = student ? student.name : (s.studentId || '');
            return [teacherName, studentName, s.score || '', s.comment || '', s.timestamp || ''];
        });

        // clear existing export rows
        await gapi.client.sheets.spreadsheets.values.clear({
            spreadsheetId: CONFIG.SPREADSHEET_ID,
            range: `${CONFIG.SHEETS.EXPORTS}!A2:E`
        });

        if (values.length > 0) {
            await gapi.client.sheets.spreadsheets.values.append({
                spreadsheetId: CONFIG.SPREADSHEET_ID,
                range: `${CONFIG.SHEETS.EXPORTS}!A2:E`,
                valueInputOption: 'USER_ENTERED',
                resource: { values }
            });
        }

        // refresh local cache
        await loadExportsData();
    } catch (err) {
        console.error('Error syncing exports sheet:', err);
        if (isSheetsPermissionError(err)) {
            showToast('您沒有權限執行此操作，請使用具有權限的帳號或聯絡管理員。', 'error');
        } else {
            showToast('更新匯出表發生錯誤', 'error');
        }
    }
}

/**
 * 載入老師資料
 */
async function loadTeachersData() {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: CONFIG.SPREADSHEET_ID,
            range: `${CONFIG.SHEETS.TEACHERS}!A2:B`,
        });
        
        const values = response.result.values || [];
        teachersData = values.map((row, index) => ({
            id: index + 1,
            name: row[0] || '',
            email: row[1] || ''
        }));
    } catch (error) {
        console.error('Error loading teachers:', error);
        if (isSheetsPermissionError(error)) {
            console.log('Detected Sheets permission error when loading teachers');
            throw error;
        }
        teachersData = [];
    }
}

/**
 * 載入學生資料
 */
async function loadStudentsData() {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: CONFIG.SPREADSHEET_ID,
            range: `${CONFIG.SHEETS.STUDENTS}!A2:B`,
        });
        
        const values = response.result.values || [];
        studentsData = values.map(row => ({
            id: row[0] || '',
            name: row[1] || ''
        }));
    } catch (error) {
        console.error('Error loading students:', error);
        if (isSheetsPermissionError(error)) throw error;
        studentsData = [];
    }
}

/**
 * 載入對應關係資料
 */
async function loadRelationsData() {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: CONFIG.SPREADSHEET_ID,
            range: `${CONFIG.SHEETS.RELATIONS}!A2:B`,
        });
        
        const values = response.result.values || [];
        relationsData = values.map((row, index) => ({
            rowIndex: index + 2, // 實際在 Sheet 中的行號（從2開始，因為第1行是標題）
            teacherEmail: row[0] || '',
            studentId: row[1] || ''
        }));
    } catch (error) {
        console.error('Error loading relations:', error);
        if (isSheetsPermissionError(error)) throw error;
        relationsData = [];
    }
}

/**
 * 載入評分資料
 */
async function loadScoresData() {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: CONFIG.SPREADSHEET_ID,
            range: `${CONFIG.SHEETS.SCORES}!A2:E`,
        });
        
        const values = response.result.values || [];
        scoresData = values.map((row, index) => ({
            rowIndex: index + 2,
            teacherEmail: row[0] || '',
            studentId: row[1] || '',
            score: row[2] || '',
            comment: row[3] || '',
            timestamp: row[4] || ''
        }));
    } catch (error) {
        console.error('Error loading scores:', error);
        if (isSheetsPermissionError(error)) throw error;
        scoresData = [];
    }
}

/**
 * 渲染學生列表
 */
function renderStudentsList() {
    elements.studentsLoading.classList.add('hidden');
    
    // 獲取當前老師可以評分的學生
    const myStudentIds = relationsData
        .filter(r => r.teacherEmail.toLowerCase() === currentUser.email.toLowerCase())
        .map(r => r.studentId);
    
    const myStudents = studentsData.filter(s => myStudentIds.includes(s.id));
    
    if (myStudents.length === 0) {
        elements.studentsEmpty.classList.remove('hidden');
        elements.studentsList.classList.add('hidden');
        return;
    }
    
    elements.studentsEmpty.classList.add('hidden');
    elements.studentsList.classList.remove('hidden');
    
    elements.studentsList.innerHTML = myStudents.map(student => {
        // 查找該學生的評分（由當前老師評的）
        const scoreRecord = scoresData.find(
            s => s.studentId === student.id && 
            s.teacherEmail.toLowerCase() === currentUser.email.toLowerCase()
        );
        
        const hasScore = scoreRecord && scoreRecord.score !== '';
        const scoreValue = hasScore ? scoreRecord.score : '未評分';
        const scoreClass = hasScore ? 'scored' : 'not-scored';
        const cardClass = hasScore ? 'student-card scored' : 'student-card';
        const btnText = hasScore ? '修改評分' : '進行評分';
        const btnClass = hasScore ? 'btn btn-secondary' : 'btn btn-primary';
        
        return `
            <div class="${cardClass}" data-student-id="${student.id}">
                <div class="student-header">
                    <div class="student-avatar">${student.name.charAt(0)}</div>
                    <div class="student-name">${student.name}</div>
                </div>
                <div class="student-score">
                    <span class="score-label">目前分數</span>
                    <span class="score-value ${scoreClass}">${scoreValue}</span>
                </div>
                <div class="student-actions">
                    <button class="${btnClass}"
                        data-student-id="${escapeAttr(student.id)}"
                        data-student-name="${escapeAttr(student.name)}"
                        data-score="${escapeAttr(scoreValue)}"
                        data-comment="${escapeAttr(scoreRecord?.comment || '')}">${btnText}</button>
                </div>
            </div>
        `;
    }).join('');

    elements.studentsList.querySelectorAll('.student-actions button').forEach(btn => {
        btn.addEventListener('click', () => {
            openScoreModal(
                btn.dataset.studentId,
                btn.dataset.studentName,
                btn.dataset.score,
                btn.dataset.comment
            );
        });
    });
}

/**
 * 渲染管理標籤頁
 */
function renderManagementTab() {
    elements.relationsLoading.classList.add('hidden');
    elements.relationsTable.classList.remove('hidden');
    
    // 填充老師下拉選單
    elements.newTeacher.innerHTML = '<option value="">請選擇老師</option>' +
        teachersData.map(t => `<option value="${t.email}">${t.name} (${t.email})</option>`).join('');
    
    // 填充學生下拉選單
    elements.newStudent.innerHTML = '<option value="">請選擇學生</option>' +
        studentsData.map(s => `<option value="${s.id}">${s.name} (${s.id})</option>`).join('');
    
    // 渲染關係表格
    renderRelationsTable();
}

/**
 * 渲染關係表格
 */
function renderRelationsTable() {
    elements.relationsTbody.innerHTML = relationsData.map(relation => {
        const teacher = teachersData.find(t => t.email.toLowerCase() === relation.teacherEmail.toLowerCase());
        const student = studentsData.find(s => s.id === relation.studentId);
        
        return `
            <tr>
                <td>${teacher ? teacher.name : relation.teacherEmail}</td>
                <td>${student ? student.name : relation.studentId}</td>
                <td>
                    <button class="btn btn-danger btn-small" onclick="deleteRelation(${relation.rowIndex})">刪除</button>
                </td>
            </tr>
        `;
    }).join('');
}

/**
 * 開啟評分 Modal
 */
function openScoreModal(studentId, studentName, currentScore, currentComment) {
    elements.modalStudentName.textContent = studentName;
    elements.scoreInput.value = currentScore !== '未評分' ? currentScore : '';
    elements.commentInput.value = currentComment;
    elements.scoreModal.dataset.studentId = studentId;
    elements.scoreModal.classList.remove('hidden');
}

/**
 * 關閉評分 Modal
 */
function closeScoreModal() {
    elements.scoreModal.classList.add('hidden');
    elements.scoreInput.value = '';
    elements.commentInput.value = '';
}

/**
 * 處理提交評分
 */
async function handleSubmitScore() {
    const studentId = elements.scoreModal.dataset.studentId;
    const score = elements.scoreInput.value;
    const comment = elements.commentInput.value;
    
    if (!score || score < 0 || score > 100) {
        showToast('請輸入有效的分數 (0-100)', 'error');
        return;
    }
    
    try {
        // 檢查是否已有評分記錄
        const existingScore = scoresData.find(
            s => s.studentId === studentId && 
            s.teacherEmail.toLowerCase() === currentUser.email.toLowerCase()
        );
        
        const timestamp = new Date().toLocaleString('zh-TW');
        
        if (existingScore) {
            // 更新現有記錄
            await gapi.client.sheets.spreadsheets.values.update({
                spreadsheetId: CONFIG.SPREADSHEET_ID,
                range: `${CONFIG.SHEETS.SCORES}!A${existingScore.rowIndex}:E${existingScore.rowIndex}`,
                valueInputOption: 'USER_ENTERED',
                resource: {
                    values: [[currentUser.email, studentId, score, comment, timestamp]]
                }
            });
            // exported sheet synchronization will be handled globally after scores reload
        } else {
            // 新增記錄
            await gapi.client.sheets.spreadsheets.values.append({
                spreadsheetId: CONFIG.SPREADSHEET_ID,
                range: `${CONFIG.SHEETS.SCORES}!A:E`,
                valueInputOption: 'USER_ENTERED',
                resource: {
                    values: [[currentUser.email, studentId, score, comment, timestamp]]
                }
            });
            // exported sheet synchronization will be handled globally after scores reload
        }
        
        showToast('評分已儲存', 'success');
        closeScoreModal();
        
        // 重新載入評分資料並更新畫面
        await loadScoresData();
        // 將評分資料同步為匯出用的姓名格式表格（保持與 scores 相同的 row order）
        await syncExportsFromScores();
        renderStudentsList();
        
    } catch (error) {
        console.error('Error saving score:', error);
        if (isSheetsPermissionError(error)) {
            showToast('您沒有權限執行此操作，請使用具有權限的帳號或聯絡管理員。', 'error');
        } else {
            showToast('儲存評分失敗', 'error');
        }
    }
}

/**
 * 處理新增對應關係
 */
async function handleAddRelation() {
    const teacherEmail = elements.newTeacher.value;
    const studentId = elements.newStudent.value;
    
    if (!teacherEmail || !studentId) {
        showToast('請選擇老師和學生', 'error');
        return;
    }
    
    // 檢查是否已存在
    const exists = relationsData.some(
        r => r.teacherEmail.toLowerCase() === teacherEmail.toLowerCase() && r.studentId === studentId
    );
    
    if (exists) {
        showToast('此對應關係已存在', 'error');
        return;
    }
    
    try {
        await gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: CONFIG.SPREADSHEET_ID,
            range: `${CONFIG.SHEETS.RELATIONS}!A:B`,
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: [[teacherEmail, studentId]]
            }
        });
        
        showToast('對應關係已新增', 'success');
        
        // 重新載入資料
        await loadRelationsData();
        renderManagementTab();
        renderStudentsList();
        
    } catch (error) {
        console.error('Error adding relation:', error);
        if (isSheetsPermissionError(error)) {
            showToast('您沒有權限執行此操作，請使用具有權限的帳號或聯絡管理員。', 'error');
        } else {
            showToast('新增對應關係失敗', 'error');
        }
    }
}

/**
 * 刪除對應關係
 */
async function deleteRelation(rowIndex) {
    if (!confirm('確定要刪除此對應關係嗎？')) {
        return;
    }
    
    try {
        // 獲取 sheet ID
        const spreadsheet = await gapi.client.sheets.spreadsheets.get({
            spreadsheetId: CONFIG.SPREADSHEET_ID
        });
        
        const relationsSheet = spreadsheet.result.sheets.find(
            s => s.properties.title === CONFIG.SHEETS.RELATIONS
        );
        
        if (!relationsSheet) {
            throw new Error('找不到對應關係工作表');
        }
        
        // 刪除該行
        await gapi.client.sheets.spreadsheets.batchUpdate({
            spreadsheetId: CONFIG.SPREADSHEET_ID,
            resource: {
                requests: [{
                    deleteDimension: {
                        range: {
                            sheetId: relationsSheet.properties.sheetId,
                            dimension: 'ROWS',
                            startIndex: rowIndex - 1,
                            endIndex: rowIndex
                        }
                    }
                }]
            }
        });
        
        showToast('對應關係已刪除', 'success');
        
        // 重新載入資料
        await loadRelationsData();
        renderManagementTab();
        renderStudentsList();
        
    } catch (error) {
        console.error('Error deleting relation:', error);
        if (isSheetsPermissionError(error)) {
            showToast('您沒有權限執行此操作，請使用具有權限的帳號或聯絡管理員。', 'error');
        } else {
            showToast('刪除對應關係失敗', 'error');
        }
    }
}

/**
 * 處理新增老師
 */
async function handleAddTeacher() {
    const name = elements.newTeacherName.value.trim();
    const email = elements.newTeacherEmail.value.trim();
    
    if (!name || !email) {
        showToast('請輸入老師姓名和 Email', 'error');
        return;
    }
    
    // 驗證 Email 格式
    if (!email.includes('@')) {
        showToast('請輸入有效的 Email', 'error');
        return;
    }
    
    // 檢查是否已存在
    const exists = teachersData.some(t => t.email.toLowerCase() === email.toLowerCase());
    if (exists) {
        showToast('此 Email 的老師已存在', 'error');
        return;
    }
    
    try {
        await gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: CONFIG.SPREADSHEET_ID,
            range: `${CONFIG.SHEETS.TEACHERS}!A:B`,
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: [[name, email]]
            }
        });
        
        showToast('老師已新增', 'success');
        elements.newTeacherName.value = '';
        elements.newTeacherEmail.value = '';
        
        // 重新載入資料
        await loadTeachersData();
        renderManagementTab();
        
    } catch (error) {
        console.error('Error adding teacher:', error);
        if (isSheetsPermissionError(error)) {
            showToast('您沒有權限執行此操作，請使用具有權限的帳號或聯絡管理員。', 'error');
        } else {
            showToast('新增老師失敗', 'error');
        }
    }
}

/**
 * 處理新增學生
 */
async function handleAddStudent() {
    const name = elements.newStudentName.value.trim();
    const studentNumber = elements.newStudentNumber.value.trim();
    
    if (!studentNumber || !name) {
        showToast('請輸入學號與學生姓名', 'error');
        return;
    }
    // 檢查學號是否已存在
    const exists = studentsData.some(s => s.id === studentNumber);
    if (exists) {
        showToast('此學號已存在，請使用不同的學號', 'error');
        return;
    }
    
    try {
        await gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: CONFIG.SPREADSHEET_ID,
            range: `${CONFIG.SHEETS.STUDENTS}!A:B`,
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: [[studentNumber, name]]
            }
        });
        
        showToast(`學生已新增 (學號: ${studentNumber})`, 'success');
        elements.newStudentName.value = '';
        elements.newStudentNumber.value = '';
        
        // 重新載入資料
        await loadStudentsData();
        renderManagementTab();
        
    } catch (error) {
        console.error('Error adding student:', error);
        if (isSheetsPermissionError(error)) {
            showToast('您沒有權限執行此操作，請使用具有權限的帳號或聯絡管理員。', 'error');
        } else {
            showToast('新增學生失敗', 'error');
        }
    }
}

/**
 * 顯示 Toast 通知
 */
function showToast(message, type = 'info') {
    elements.toast.textContent = message;
    elements.toast.className = `toast ${type}`;
    elements.toast.classList.remove('hidden');
    
    setTimeout(() => {
        elements.toast.classList.add('hidden');
    }, 3000);
}

// 將需要從 HTML 呼叫的函數暴露到全域
window.deleteRelation = deleteRelation;
