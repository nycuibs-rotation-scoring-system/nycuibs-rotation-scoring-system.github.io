# 學生評分系統

這是一個純前端的學生評分網站，使用 Google Sheets 作為資料儲存媒介。

## 功能特點

- ✅ Google OAuth 登入驗證
- ✅ 老師權限控管（只有在名單中的老師可以使用）
- ✅ 老師可評分指定的學生
- ✅ 即時顯示學生的評分狀態
- ✅ 名單設定（新增、刪除學生對應）
- ✅ 支援新增老師和學生
- ✅ 所有資料儲存在 Google Sheets

## 設定步驟

### 1. 建立 Google Sheets

1. 前往 [Google Sheets](https://sheets.google.com) 建立新的試算表
1. 建立以下 5 個工作表（Sheet）：

#### 工作表 1：老師名單

| 姓名   | Email              |
| :----- | :----------------- |
| 王老師 | `wang@gmail.com`   |
| 李老師 | `lee@gmail.com`    |

#### 工作表 2：學生名單

| 學號      | 姓名 |
| :-------- | :--- |
| 300113000 | 小明 |
| 300113001 | 小華 |
| 300113002 | 小美 |

#### 工作表 3：對應關係

| 老師 Email         | 學號      |
| :----------------- | :-------- |
| `wang@gmail.com`   | 300113000 |
| `wang@gmail.com`   | 300113001 |
| `lee@gmail.com`    | 300113001 |
| `lee@gmail.com`    | 300113002 |

#### 工作表 4：評分紀錄

| 老師 Email       | 學號 | 分數 | 備註 | 評分時間 |
| :--------------- | :--- | :--- | :--- | :------- |
| （系統自動填入） |      |      |      |          |

#### 工作表 5：評分匯出（匯出用，與評分紀錄一一對應）

| 老師姓名 | 學生姓名 | 分數 | 備註     | 評分時間   |
| :------- | :------- | :--- | :------- | :--------- |
| 王老師   | 小明     | 85   | 表現良好 | 2024-01-01 |

說明：`評分匯出` 工作表的每一列對應 `評分紀錄` 的相同行（從第 2 行開始），系統會在新增或修改評分後重新生成 `評分匯出` 的內容，保證行序與 `評分紀錄` 一致。此表僅包含姓名/分數/評語/時間，方便直接匯出給非技術使用者。

1. 將試算表分享給所有老師的 Gmail，設定為「編輯者」權限

### 2. 設定 Google Cloud 專案

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
1. 建立新專案或選擇現有專案
1. 啟用以下 API：
   - Google Sheets API
   - Google Identity Services

1. 建立 OAuth 2.0 憑證：
   - 前往「API 和服務」>「憑證」
   - 點擊「建立憑證」>「OAuth 用戶端 ID」
   - 應用程式類型選擇「網頁應用程式」
   - 新增授權的 JavaScript 來源（例如：`http://localhost:5500` 或您的網站網址）
   - 新增授權的重新導向 URI（與上面相同）
   - 複製產生的「用戶端 ID」

1. 設定 OAuth 同意畫面：
   - 前往「API 和服務」>「OAuth 同意畫面」
   - 選擇「外部」（除非您有 Google Workspace）
   - 填寫應用程式名稱、使用者支援電子郵件等必填欄位
   - 新增範圍：`openid`, `email`, `profile` 以及 `https://www.googleapis.com/auth/spreadsheets`
   - 如果處於測試模式，新增測試使用者（老師的 Gmail）

   注意：系統在登入成功階段會自動檢查您對目標 Google Sheets 的讀取/寫入存取；若該帳號沒有必要權限，系統會顯示明確錯誤並要求您改用有權限的帳號或聯絡擁有者授權。

### 3. 設定專案配置

建議不要直接修改 `config.js`，而是複製 `env.example.js` 為 `env.js`，再填入部署用設定。

1. 複製 `env.example.js` 為 `env.js`
1. 依實際環境修改 `env.js` 內容
1. 確認 `index.html` 中有先載入 `env.js`，再載入 `config.js`

範例：

```javascript
window.APP_CONFIG = {
   GOOGLE_CLIENT_ID: 'YOUR_CLIENT_ID.apps.googleusercontent.com',
   SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID',
   LEGAL_CONTACT: '系統管理者名稱'
};
```

說明：

- `GOOGLE_CLIENT_ID`：Google OAuth 2.0 Web client ID
- `SPREADSHEET_ID`：Google Sheets 網址中的試算表 ID
- `LEGAL_CONTACT`：法律頁面顯示的聯絡名稱

### 4. 啟動網站

由於使用了 Google OAuth，需要透過 HTTP 伺服器存取網頁：

#### 方法 1：使用 VS Code Live Server

- 安裝並啟用 Live Server 擴充套件
- 在 `index.html` 上按右鍵，選擇 `Open with Live Server`
- 確認開啟的網址已加入 Google OAuth 的授權來源，例如 `http://127.0.0.1:5500` 或 `http://localhost:5500`

#### 方法 2：使用 Python

```bash
# Python 3
python -m http.server 5500

# 然後開啟 http://localhost:5500
```

#### 方法 3：使用 Node.js

```bash
npx serve -p 5500

# 然後開啟 http://localhost:5500
```

## 使用說明

### 老師操作流程

1. **登入**：點擊「使用 Google 登入」，使用您的 Gmail 帳號登入
2. **評分**：
   - 登入後會看到您可以評分的學生清單
   - 每張學生卡片會顯示目前的評分狀態
   - 點擊「進行評分」或「修改評分」按鈕
   - 輸入分數（0-100）和備註
   - 點擊「提交評分」儲存

3. **管理對應關係**（如需要）：
   - 切換到「名單設定」標籤頁
   - 可以新增老師、學生
   - 可以整理自己的評分名單，新增或刪除學生對應

## 檔案結構

```text
├── index.html      # 主頁面
├── styles.css      # 樣式表
├── config.js       # 配置檔（需要修改）
├── app.js          # 主應用程式邏輯
└── README.md       # 說明文件
```

## 注意事項

1. **權限控管**：只有在「老師名單」工作表中的 Email 才能登入使用系統
2. **評分限制**：老師只能評分「對應關係」工作表中指派給他的學生
3. **資料安全**：所有資料儲存在您的私有 Google Sheets 中
4. **瀏覽器相容性**：建議使用 Chrome、Firefox、Edge 等現代瀏覽器

### 權限錯誤處理

- 若登入的 Google 帳號沒有該試算表的讀取或編輯權限，系統會顯示明確錯誤提示並要求您使用有權限的帳號或向擁有者請求「編輯者」權限。
- 在某些寫入操作失敗（例如新增學生、評分、刪除關係）時，系統會檢測到權限錯誤並顯示友善錯誤訊息，避免不清楚的錯誤代碼。

### 登入狀態保存

- 本專案會在登入成功後將授權 token 與使用者資訊儲存在瀏覽器的 localStorage（key 為 `grading_app_session_v1`），以便在重新整理或重新載入頁面時保持登入狀態。
- 注意：access token 會過期；若 token 已過期或被撤銷，網站會自動回到登入畫面並需重新登入。
- 如需強制重新授權，請在登入時選擇「同意」或使用開發工具清除 localStorage 後重新登入。
- 安全性提醒：將 access token 儲存在 localStorage 會有風險（XSS 攻擊可能讀取 token）。若需要更高安全性，請使用受保護的後端或更嚴格的前端策略。

### 法律頁面聯絡資訊

- 可在 `env.js` 或部署時注入的 `window.APP_CONFIG` 中設定 `LEGAL_CONTACT`。
- `privacy-policy.html` 與 `terms-of-service.html` 會自動讀取這個值；若未提供，頁面會退回顯示「本系統管理者」。

### GitHub Actions 部署注入

- 現有的 [deploy workflow](.github/workflows/deploy.yml) 會在部署時產生 `env.js`。
- 請在 GitHub repository secrets 中設定以下三個值：`GOOGLE_CLIENT_ID`、`SPREADSHEET_ID`、`LEGAL_CONTACT`。
- workflow 產生的格式如下：

```javascript
window.APP_CONFIG = {
   GOOGLE_CLIENT_ID: 'your-client-id.apps.googleusercontent.com',
   SPREADSHEET_ID: 'your-spreadsheet-id',
   LEGAL_CONTACT: '系統管理者名稱'
};
```

## 常見問題

### Q: 登入時出現「此帳號不在老師名單中」？

A: 請確認您使用的 Gmail 是否已加入 Google Sheets 的「老師名單」工作表中。

### Q: 無法存取 Google Sheets？

A: 請確認：

1. Google Sheets API 已啟用
2. OAuth 同意畫面已設定
3. 測試使用者已加入（如果 App 處於測試模式）

### Q: 評分無法儲存？

A: 請確認您的 Gmail 帳號有該 Google Sheets 的編輯權限。

## 技術細節

- 純前端架構，無後端伺服器
- 使用 Google Identity Services (GIS) 進行 OAuth 2.0 認證
- 使用 Google Sheets API v4 進行資料讀寫
- 響應式設計，支援桌面和行動裝置
