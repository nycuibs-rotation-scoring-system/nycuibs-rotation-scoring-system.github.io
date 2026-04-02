/**
 * 配置文件
 * 請根據您的 Google Cloud 專案和 Google Sheets 設定修改以下配置
 */

// Read configuration from a preloaded global (recommended for static deployment)
// or from process.env (when bundling). The flow is:
// 1) Create a non-committed `env.js` that sets `window.APP_CONFIG = { GOOGLE_CLIENT_ID: '...', SPREADSHEET_ID: '...' }`
// 2) Include that `env.js` BEFORE config.js in your HTML (index.html already has a slot)
// 3) If window.APP_CONFIG is not present, fallback to process.env or the default literal values below.

const _ENV = (typeof window !== 'undefined' && window.APP_CONFIG) || (typeof process !== 'undefined' && process.env) || {};

const CONFIG = {
    // Google OAuth 2.0 Client ID
    // 請到 Google Cloud Console 創建 OAuth 2.0 憑證
    // https://console.cloud.google.com/apis/credentials
    // GOOGLE_CLIENT_ID: 'YOUR_GOOGLE_CLIENT_ID.apps.googleusercontent.com',
    GOOGLE_CLIENT_ID: _ENV.GOOGLE_CLIENT_ID || '',

    // Google Sheets ID
    // 從 Google Sheets URL 中獲取: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit
    // SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID',
    SPREADSHEET_ID: _ENV.SPREADSHEET_ID || '',

    // 法律頁面聯絡資訊
    LEGAL_CONTACT: _ENV.LEGAL_CONTACT || '',
    // SPREADSHEET_ID: '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',


    // Google API Key (用於讀取公開資料，但我們使用 OAuth 所以這個可選)
    // API_KEY: 'YOUR_API_KEY',

    // Sheet 名稱配置
    SHEETS: {
        TEACHERS: '老師名單',      // 老師名單工作表名稱
        STUDENTS: '學生名單',      // 學生名單工作表名稱
        RELATIONS: '對應關係',     // 老師-學生對應關係工作表名稱
        SCORES: '評分紀錄'         // 評分紀錄工作表名稱
        ,EXPORTS: '評分匯出'      // 匯出用：老師姓名/學生姓名 的紀錄（同步於評分紀錄）
    },

    // Google API scopes
    // 必須包含 openid email profile，才能在 userinfo endpoint 拿到 email 與 profile
    // 同時保留 sheets scope 用於讀寫試算表
    SCOPES: 'openid email profile https://www.googleapis.com/auth/spreadsheets'
};

/**
 * Google Sheets 結構說明：
 * 
 * 1. 老師名單 (Teachers):
 *    | A (姓名) | B (Email) |
 *    |----------|-----------|
 *    | 王老師   | wang@gmail.com |
 *    | 李老師   | lee@gmail.com  |
 * 
 * 2. 學生名單 (Students):
 *    | A (學號) | B (姓名) |
 *    |------------|----------|
 *    | 300113000  | 小明     |
 *    | 300113001  | 小華     |
 * 
 * 3. 對應關係 (Relations):
 *    | A (老師Email) | B (學號) |
 *    |---------------|------------|
 *    | wang@gmail.com | 300113000      |
 *    | wang@gmail.com | 300113001      |
 *    | lee@gmail.com  | 300113001      |
 * 
 * 4. 評分紀錄 (Scores):
 *    | A (老師Email) | B (學號) | C (分數) | D (備註) | E (評分時間) |
 *    |---------------|------------|----------|----------|--------------|
 *    | wang@gmail.com | 300113000      | 85       | 表現良好 | 2024-01-01   |
 */
