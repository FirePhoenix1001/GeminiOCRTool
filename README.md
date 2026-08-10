# GeminiOCR - 100% 純前端網頁數學 OCR 與詳解工具 🌻

這是一個採用 **「100% 純靜態網頁（純前端 JavaScript）」** 架構重構的數學公式 OCR 轉錄與詳解生成工具。
所有運算都在您的瀏覽器端安全、私密、快速地運行，完全擺脫本機 Python、C# 與伺服器環境依賴。

---

## 🚀 快速開始使用

### 🟢 使用者（雙擊即用）
1. 在檔案總管中，**直接雙擊專案根目錄下的 `index.html`**，它將會以您的預設瀏覽器（推薦使用 Chrome 或 Edge）開啟。
2. 切換到 **金鑰與模型設定** 頁籤，輸入您的 Google Gemini 或 OpenAI API Key 並點擊儲存，即可立即開始使用！

### 💻 開發者（現代前端開發工作流）
本專案已全面導入現代前端工程化配置（Vite + ESM 模組化開發 + 自動打包單一 HTML）：
1. 確保您的系統已安裝 Node.js 與 npm。
2. 在專案根目錄執行以下指令安裝依賴套件：
   ```bash
   npm install
   ```
3. 開啟本機熱重載開發伺服器（Live Reload）：
   ```bash
   npm run dev
   ```
   瀏覽器會自動開啟 `http://localhost:5173` 進行即時修改與測試。
4. 編譯打包發布：
   ```bash
   npm run build
   ```
   此指令會自動將 `src/web/` 的 ESM 模組編譯並融入為單一的自包含 `dist/index.html`，並透過腳本複製回最外層的 `index.html`。

---

## 💎 特色功能與技術亮點

*   **100% 瀏覽器本地運算**：使用 Mozilla 的 `PDF.js` 將 PDF 檔案渲染並以 Base64 PNG 圖片提取；使用 `docx.js` 在瀏覽器內動態打包與產出 Times New Roman、標楷體混合排版之 Word 檔案。
*   **指數型退避重試機制 (Exponential Backoff)**：針對 API 限流或伺服器 busy (如 `high demand` / `429` 錯誤)，前台會自動啟動指數延遲退避重試，最大程度確保批次轉換不中斷。
*   **多組金鑰輪替管理**：支援填入多個 API Key。當某一組 Key 的每日限額用盡或連線異常時，前台將自動切換下一組 Key 繼續執行。
*   **Word XML 排版優化 (JSZip)**：支援將現有的 `.docx` 檔案拖入，利用 `JSZip` 於前端解壓並直接替換優化 `word/document.xml` 內容，實現減號 `-` 與 `−` 互換以及上/下標結構轉換。
*   **本地金鑰儲存隱私**：所有 API Key、Prompt 規則與自定義模型，皆以加密或本機快取形式保存在瀏覽器的 `localStorage` 中，100% 保障您的隱私安全。

---

## ⚙️ 專案目錄結構

*   **`C:\PythonProgram\GeminiOCRTool\`** (專案根目錄)
    *   `index.html` - 自動打包後的最終自包含 HTML 檔案（雙擊即可直接運行，無需任何環境）。
    *   `prompt.txt` - OCR 數學轉錄專用提示詞規則檔。
    *   `prompt_explain.txt` - 詳解生成專用提示詞規則檔。
    *   `package.json` - 前端套件依賴與 npm 打包腳本。
    *   `vite.config.js` - Vite 編譯與單一 HTML 插件配置。
    *   `copy-build.js` - 自動複製打包結果至根目錄的 Node.js 腳本。
    *   `README.md` - 本說明書。
    *   **`src/`** - 原始碼目錄
        *   **`web/`** - 前端 ESM 核心程式開發目錄
            *   `index.html` - 網頁 UI 結構範本。
            *   **`css/`**
                *   `style.css` - 磨砂玻璃樣式表。
            *   **`js/`**
                *   `app.js` - 主控協調模組 (Orchestrator)。
                *   `api.js` - API 通訊、Key 輪替與退避重試模組。
                *   `pdf.js` - PDF.js 文件解析與渲染模組。
                *   `word.js` - Word 生成與 XML 優化重寫模組。