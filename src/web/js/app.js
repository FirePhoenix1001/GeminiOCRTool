// app.js - GeminiOCR Orchestrator / UI Controller

import { callAIApi, setLogger } from './api.js';
import { convertPdfPageToPngBase64, readImageAsBase64, getPdfPageCount } from './pdf.js';
import { buildWordDocument, triggerTxtDownload, optimizeWordFile } from './word.js';
import { VERSION, DEFAULT_MODEL } from './config.js';

// Default rules (Prompts) from original project
const DEFAULT_OCR_RULE = `<role>
你是一位追求「100% 忠實還原」的 OCR 數學數位化專家，擅長以「純文字為主、LaTeX 為輔」的方式排版。
</role>

<rules>
1. **LaTeX 使用的絕對禁令（環境繼承原則）：**
   - **唯一例外：** **只有在出現「分數 \\frac{}{}」、「根號 \\sqrt{}{}」、「聯立方程式（如 \\begin{cases}）」及「矩陣（如 \\begin{matrix}）」時，才允許使用 $...$。**
   - **環境繼承：** 一旦進入 $ 內部，所有內容（含運算符、上下標、變數、矩陣元素）必須統一採用 LaTeX 寫法（如 \\times, n_{1}, \\text{matrix} 語法）。

2. **純文字優先與簡化標記（最高準則）：**
   - **一般變數：** 所有的變數、數字、希臘字母、邏輯符號、比較符號，只要在 $ 之外，一律採純文字，嚴禁包進 $ 內。
   - **上下標處理：** 在 $ 之外時，直接使用 \`_{}\` 表示下標，\`^{}\` 表示上標（例如：L_{1}, x^{2}, (y-2)^{2}）。
   - **範例：** 使用 x≈y±z 而不是 $x$≈$y$±$z$。

3. **符號清單（限純文字顯示）：**
   - **關係：** <, >, =, ≤, ≥, ≠, ≈, ±, ⊥, °。
   - **邏輯：** ∵, ∴, ⇒, ⇔, ∈。
   - **運算：** ×, -, ·, π, …, ∠, $\\overline{AB}$ (上劃線仍可用 $)。
   - **希臘/序號：** θ, α, β, Γ, ①, ②。
   - **函數：** log, sin, cos, tan。
   - **是非 :  ** ◯, ✕。

4. **排版要求：**
   - **圖片標記：** 若有輔助圖片，請換行註記 *********[缺圖]*********。
   - **結構保持：** 保持原始段落與行序，題號格式依原樣，結尾「選( )」照原樣。
   - **禁令：** 忽略最下方頁碼；不使用任何 Markdown 格式語法。
</rules>

<examples>
以下為「純文字優先」的正確排版範例：

- **變數與比較（不包 $）：**
  正確：x≈y±z
  正確：x∈A ⇒ x≠①

- **幾何與邏輯：**
  正確：∵ ∠α=∠β ∴ △ABC 為等腰三角形。
  正確：$\\overline{AB}$=\\overline{AC}$

- **矩陣處理（必須使用 LaTeX）：**
  正確：若 A = $\\left[ \\begin{matrix} -2 & 4 \\\\ 1 & 3 \\\\ 7 & -3 \\end{matrix} \\right]$ , 則 2A = $\\left[ \\begin{matrix} -4 & 4 \\\\ 2 & 3 \\\\ 14 & -3 \\end{matrix} \\right]$

- **必須使用 LaTeX 的情境（分數、根號、聯立）：**
  正確：$\\frac{1}{2}$+$\\frac{\\sqrt{3}}{2}$ = $\\frac{1+\\sqrt{3}}{2}$
  正確：(A) $\\begin{cases} a+2\\theta\\beta>0 \\\\ 7a-24b\\theta\\beta+22>0 \\end{cases}$

- **函數與運算：**
  正確：log 2 + sin θ = 1
  正確：3 × 2 = 6

- **一般上下標（使用純文字標記）：**
  正確：L_{1}: 2x-y+a=0
  正確：(x+3)^{2}+(y-2)^{2}=60

- **複合環境（環境繼承）：**
  正確：$\\frac{\\pi n_{1}}{\\pi}$ = n_{1}
  （解釋：因為在分數內，所以 n_{1} 必須進入 $ 環境；但等號後的結果回歸純文字標記。）
</examples>

<task>
根據上述「純文字標記優先」規則，將提供圖片內容轉錄為文字。
你的核心任務是將圖片中的數學內容「完全相符」地轉錄為文字，嚴禁任何擅自的簡化、省略 or 美化。
</task>`;

const DEFAULT_EXPLAIN_RULE = `<role>
你是一位極其專業且耐心的「數學詳解專家」，擅長以「純文字為主、LaTeX 為輔」的方式，為數學題目提供步驟詳盡、邏輯嚴密的解答。
你的目標是：不僅給出正確答案，更要讓學生看懂解題的每一個轉折。
</role>

<rules>
1. **LaTeX 使用的絕對禁令（環境繼承原則）：**
   - **唯一例外：** ** 只有在出現「分數 \\frac{}{}」、「根號 \\sqrt{}{}」及「**聯立方程式（如 \\begin{cases}）**」時，才允許使用 $...$。
   - **環境繼承：** 一旦進入 $ 內部，所有內容（含運算符、上下標、變數）必須統一採用 LaTeX 寫法（如 \\times, n_{1}）。

2. **純文字優先與簡化標記（最高準則）：**
   - **一般變數：** 所有的變數、數字、希臘字母、邏輯符號、比較符號，只要在 $ 之外，一律採純文字，嚴禁包進 $ 內。
   - **上下標處理：** 在 $ 之外時，直接使用 \`_{}\` 表示下標，\`^{}\` 表示上標（例如：L_{1}, x^{2}, (y−2)^{2}）。
   - **範例：** 使用 x≈y±z 而不是 $x$≈$y$±$z$。

3. **符號清單（限純文字顯示）：**
   - **關係：** <, >, =, ≤, ≥, ≠, ≈, ±。
   - **邏輯：** ∵, ∴, ⇒, ⇔, ∈。
   - **運算：** ×, ·, π, …, ∠, $\\overline{AB}$ (上劃線仍可用 $)。
   - **希臘/序號：** θ, α, β, Γ, ①, ②。
   - **函數：** log, sin, cos, tan。

4. **解題結構：**
   - **題目：** 先簡短摘要題目內容。
   - **詳解：** 分步驟說明解法，推導邏輯需流暢且明確。
   - **答案：** 在最後明確標註最終答案。
   - **禁令：** 不使用 any Markdown 格式語法。
</rules>

<examples>
以下為符合「純文字優先」規則的詳解排版範例：

- **題目摘要：** 若 x+y=5 且 xy=6，求 x^{2}+y^{2} 之值。
- **詳解：** 
  ∵ (x+y)^{2} = x^{2}+2xy+y^{2}
  ∴ 5^{2} = x^{2}+2(6)+y^{2}
  25 = x^{2}+12+y^{2}
  x^{2}+y^{2} = 25−12 = 13
- **答案：** 13
</examples>

<task>
根據上述「純文字標記優先」規則與「詳解專家」角色，分析圖片中的數學題目並生成詳解。
請確保推導過程嚴謹，格式美觀且符合上述 LaTeX 禁令與純文字優先原則。
</task>`;

document.addEventListener('DOMContentLoaded', () => {
    // Arrays to hold files
    let ocrFiles = [];
    let explainFiles = [];
    let wordFiles = [];

    let isProcessing = false;
    let stopRequested = false;

    // Load UI elements
    const tabs = document.querySelectorAll('.nav-tab');
    const tabPanes = document.querySelectorAll('.tab-pane');
    const themeToggleBtn = document.getElementById('theme-toggle-btn');
    const apiStatusBadge = document.getElementById('api-status-badge');
    const statusAlertBanner = document.getElementById('status-alert-banner');
    const terminalConsole = document.getElementById('terminal-console');
    const clearLogsBtn = document.getElementById('clear-logs-btn');

    // Progress elements
    const progressContainer = document.getElementById('progress-container');
    const progressTaskName = document.getElementById('progress-task-name');
    const progressPercentage = document.getElementById('progress-percentage');
    const progressBarFill = document.getElementById('progress-bar-fill');

    // Settings elements
    const keysListContainer = document.getElementById('keys-list-container');
    const addKeyBtn = document.getElementById('add-key-btn');
    const modelSelect = document.getElementById('model-select');
    const customModelWrapper = document.getElementById('custom-model-wrapper');
    const customModelNameInput = document.getElementById('custom-model-name');
    const settingsIgnoreHandwriting = document.getElementById('settings-ignore-handwriting');
    const saveSettingsBtn = document.getElementById('save-settings-btn');

    let currentApiKeys = [];

    // Modal elements
    const ruleModal = document.getElementById('rule-modal');
    const ruleModalClose = document.getElementById('rule-modal-close');
    const ruleModalTitle = document.getElementById('rule-modal-title');
    const ruleModalTextarea = document.getElementById('rule-modal-textarea');
    const ruleModalResetBtn = document.getElementById('rule-modal-reset-btn');
    const ruleModalAiBtn = document.getElementById('rule-modal-ai-btn');
    const ruleModalSaveBtn = document.getElementById('rule-modal-save-btn');

    const aiRefineModal = document.getElementById('ai-refine-modal');
    const aiRefineModalClose = document.getElementById('ai-refine-modal-close');
    const refineBadOutput = document.getElementById('refine-bad-output');
    const refineGoodOutput = document.getElementById('refine-good-output');
    const refineResultOutput = document.getElementById('refine-result-output');
    const refineGenerateBtn = document.getElementById('refine-generate-btn');
    const refineApplyBtn = document.getElementById('refine-apply-btn');

    // Tab state
    let activeRuleType = 'ocr'; // 'ocr' or 'explain'

    /* ==========================================
       1. Logging & Alerts
       ========================================== */

    function logMsg(text, type = '') {
        const line = document.createElement('div');
        line.className = 'log-line';
        if (type === 'error') line.classList.add('error-line');
        if (type === 'system' || text.startsWith('[SYSTEM]')) line.classList.add('system-line');

        line.textContent = text;
        terminalConsole.appendChild(line);
        terminalConsole.scrollTop = terminalConsole.scrollHeight;
    }

    // Register our UI logger with the API module
    setLogger(logMsg);
    logMsg(`[SYSTEM] 當前版本：${VERSION}`);

    clearLogsBtn.addEventListener('click', () => {
        terminalConsole.innerHTML = '<div class="log-line system-line">[SYSTEM] 控制台日誌已清除。🌻</div>';
    });

    function showToast(message, type = 'success') {
        const toast = document.createElement('div');
        toast.className = `custom-toast toast-${type}`;
        
        let iconClass = 'fa-check-circle';
        if (type === 'warning') iconClass = 'fa-exclamation-circle';
        if (type === 'error') iconClass = 'fa-times-circle';
        if (type === 'info') iconClass = 'fa-circle-info';
        
        toast.innerHTML = `
            <i class="fa-solid ${iconClass}"></i>
            <span>${message}</span>
        `;
        
        if (!document.getElementById('toast-style-tag')) {
            const style = document.createElement('style');
            style.id = 'toast-style-tag';
            style.innerHTML = `
                .custom-toast {
                    position: fixed;
                    top: 20px;
                    right: 20px;
                    background: rgba(30, 24, 46, 0.95);
                    backdrop-filter: blur(8px);
                    border: 1px solid var(--border-color-active);
                    color: #fff;
                    padding: 14px 24px;
                    border-radius: 12px;
                    z-index: 99999;
                    display: flex;
                    align-items: center;
                    gap: 12px;
                    box-shadow: 0 10px 30px rgba(0,0,0,0.5);
                    transform: translateY(-20px);
                    opacity: 0;
                    transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275);
                }
                .custom-toast.show {
                    transform: translateY(0);
                    opacity: 1;
                }
                .toast-success i { color: #2cc985; }
                .toast-warning i { color: #ffb300; }
                .toast-error i { color: #ff5555; }
                .toast-info i { color: #3b8ed0; }
            `;
            document.head.appendChild(style);
        }
        
        document.body.appendChild(toast);
        setTimeout(() => toast.classList.add('show'), 50);
        
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 300);
        }, 4000);
    }

    /* ==========================================
       2. Navigation Tabs & Theme Toggle
       ========================================== */

    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            tabPanes.forEach(p => p.classList.remove('active'));

            tab.classList.add('active');
            const targetPane = document.getElementById(tab.dataset.tab);
            targetPane.classList.add('active');
        });
    });

    themeToggleBtn.addEventListener('click', () => {
        const isDark = document.body.classList.contains('theme-dark');
        if (isDark) {
            document.body.classList.remove('theme-dark');
            document.body.classList.add('theme-light');
            themeToggleBtn.innerHTML = '<i class="fa-solid fa-sun"></i> <span>淺色模式</span>';
            localStorage.setItem('geminiocr-theme', 'light');
        } else {
            document.body.classList.remove('theme-light');
            document.body.classList.add('theme-dark');
            themeToggleBtn.innerHTML = '<i class="fa-solid fa-moon"></i> <span>深色模式</span>';
            localStorage.setItem('geminiocr-theme', 'dark');
        }
    });

    if (localStorage.getItem('geminiocr-theme') === 'light') {
        document.body.classList.remove('theme-dark');
        document.body.classList.add('theme-light');
        themeToggleBtn.innerHTML = '<i class="fa-solid fa-sun"></i> <span>淺色模式</span>';
    }

    document.querySelectorAll('.toggle-password-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const input = btn.previousElementSibling;
            const icon = btn.querySelector('i');
            if (input.type === 'password') {
                input.type = 'text';
                icon.className = 'fa-solid fa-eye-slash';
            } else {
                input.type = 'password';
                icon.className = 'fa-solid fa-eye';
            }
        });
    });

    /* ==========================================
       3. Settings Management
       ========================================== */

    modelSelect.addEventListener('change', () => {
        if (modelSelect.value === 'custom') {
            customModelWrapper.style.display = 'block';
        } else {
            customModelWrapper.style.display = 'none';
        }
    });

    // Dynamic Keys Rendering Function
    function renderKeysList() {
        keysListContainer.innerHTML = '';
        if (currentApiKeys.length === 0) {
            keysListContainer.innerHTML = '<div class="empty-list-text" style="padding: 10px 0;">目前無設定任何金鑰，請點擊新增按鈕</div>';
            return;
        }

        currentApiKeys.forEach((item, index) => {
            const row = document.createElement('div');
            row.className = 'key-row';
            row.dataset.index = index;

            row.innerHTML = `
                <select class="key-type-select">
                    <option value="gemini" ${item.type === 'gemini' ? 'selected' : ''}>Gemini</option>
                    <option value="openai" ${item.type === 'openai' ? 'selected' : ''}>OpenAI</option>
                </select>
                <div class="password-input-wrapper">
                    <input type="password" class="key-value-input" value="${item.key}" placeholder="請輸入 API Key">
                    <button class="toggle-password-btn" type="button"><i class="fa-solid fa-eye"></i></button>
                </div>
                <button type="button" class="delete-key-row-btn" title="刪除金鑰"><i class="fa-solid fa-trash-can"></i></button>
            `;

            // Bind change on type select
            row.querySelector('.key-type-select').addEventListener('change', (e) => {
                currentApiKeys[index].type = e.target.value;
            });

            // Bind input on key value
            row.querySelector('.key-value-input').addEventListener('input', (e) => {
                currentApiKeys[index].key = e.target.value.trim();
            });

            // Bind toggle visibility
            row.querySelector('.toggle-password-btn').addEventListener('click', () => {
                const input = row.querySelector('.key-value-input');
                const icon = row.querySelector('.toggle-password-btn i');
                if (input.type === 'password') {
                    input.type = 'text';
                    icon.className = 'fa-solid fa-eye-slash';
                } else {
                    input.type = 'password';
                    icon.className = 'fa-solid fa-eye';
                }
            });

            // Bind delete row
            row.querySelector('.delete-key-row-btn').addEventListener('click', () => {
                currentApiKeys.splice(index, 1);
                renderKeysList();
            });

            keysListContainer.appendChild(row);
        });
    }

    addKeyBtn.addEventListener('click', () => {
        currentApiKeys.push({ type: 'gemini', key: '' });
        renderKeysList();
    });

    function loadSettings() {
        // Load API Keys from geminiocr-api-keys
        let savedKeys = [];
        try {
            const keysStr = localStorage.getItem('geminiocr-api-keys');
            if (keysStr) {
                savedKeys = JSON.parse(keysStr);
            } else {
                // Backward compatibility: Migrate old keys
                const primaryGemini = localStorage.getItem('gemini-key') || '';
                const primaryOpenai = localStorage.getItem('openai-key') || '';
                if (primaryGemini) {
                    savedKeys.push({ type: 'gemini', key: primaryGemini });
                }
                if (primaryOpenai) {
                    savedKeys.push({ type: 'openai', key: primaryOpenai });
                }
                // Migrate old multi keys if enabled
                const useMulti = localStorage.getItem('use-multi') === 'true';
                const multiKeysText = localStorage.getItem('multi-keys') || '';
                if (useMulti && multiKeysText) {
                    const lines = multiKeysText.split('\n');
                    lines.forEach(line => {
                        line = line.trim();
                        if (!line) return;
                        let k = line;
                        let t = 'gemini';
                        if (line.includes(':')) {
                            const parts = line.split(':', 2);
                            k = parts[1].trim();
                        }
                        t = k.startsWith('sk-') ? 'openai' : 'gemini';
                        if (!savedKeys.some(item => item.key === k)) {
                            savedKeys.push({ type: t, key: k });
                        }
                    });
                }
            }
        } catch (e) {
            console.error('Migration or parse of API keys failed:', e);
        }

        if (savedKeys.length === 0) {
            savedKeys.push({ type: 'gemini', key: '' });
        }

        currentApiKeys = savedKeys;
        renderKeysList();

        modelSelect.value = localStorage.getItem('model-name') || DEFAULT_MODEL;
        customModelNameInput.value = localStorage.getItem('custom-model') || '';
        settingsIgnoreHandwriting.checked = localStorage.getItem('ignore-handwriting') === 'true';

        modelSelect.dispatchEvent(new Event('change'));
        updateApiKeyStatus();
    }

    function updateApiKeyStatus() {
        const hasKeys = currentApiKeys.some(item => item.key.trim() !== '');

        if (hasKeys) {
            apiStatusBadge.className = 'status-badge good';
            apiStatusBadge.textContent = '🟢 ON';
            apiStatusBadge.title = 'API 金鑰已配置';
            statusAlertBanner.classList.add('hide');
        } else {
            apiStatusBadge.className = 'status-badge bad';
            apiStatusBadge.textContent = '🔴 OFF';
            apiStatusBadge.title = '尚未設定 API 金鑰';
            statusAlertBanner.classList.remove('hide');
        }
    }

    saveSettingsBtn.addEventListener('click', () => {
        // Filter out empty rows when saving
        const filteredKeys = currentApiKeys.filter(item => item.key.trim() !== '');
        
        // Save new format
        localStorage.setItem('geminiocr-api-keys', JSON.stringify(filteredKeys));
        
        // Backward compatibility: Save first valid keys to legacy format
        const firstGemini = filteredKeys.find(item => item.type === 'gemini');
        const firstOpenai = filteredKeys.find(item => item.type === 'openai');
        localStorage.setItem('gemini-key', firstGemini ? firstGemini.key.trim() : '');
        localStorage.setItem('openai-key', firstOpenai ? firstOpenai.key.trim() : '');

        localStorage.setItem('model-name', modelSelect.value);
        localStorage.setItem('custom-model', customModelNameInput.value.trim());
        localStorage.setItem('ignore-handwriting', settingsIgnoreHandwriting.checked);

        // Update key rotation index in memory
        currentApiKeys = filteredKeys.length > 0 ? filteredKeys : [{ type: 'gemini', key: '' }];
        renderKeysList();

        updateApiKeyStatus();
        showToast('設定已儲存成功！ 🌻', 'success');
        logMsg('[SYSTEM] 系統核心金鑰與模型設定已更新並儲存。');
    });

    loadSettings();

    /* ==========================================
       4. File List & Dropzone Binding
       ========================================== */

    function registerDropZone(dropZoneId, inputId, filesArray, listId, countId, type) {
        const dropZone = document.getElementById(dropZoneId);
        const input = document.getElementById(inputId);

        dropZone.addEventListener('click', () => input.click());

        input.addEventListener('change', (e) => {
            handleSelectedFiles(e.target.files, filesArray, listId, countId, type);
            input.value = '';
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            dropZone.addEventListener(eventName, (e) => {
                e.preventDefault();
                dropZone.classList.add('dragover');
            }, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, (e) => {
                e.preventDefault();
                dropZone.classList.remove('dragover');
            }, false);
        });

        dropZone.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            handleSelectedFiles(dt.files, filesArray, listId, countId, type);
        });
    }

    registerDropZone('ocr-drop-zone', 'ocr-file-input', ocrFiles, 'ocr-file-list', 'ocr-file-count', 'ocr');
    registerDropZone('explain-drop-zone', 'explain-file-input', explainFiles, 'explain-file-list', 'explain-file-count', 'explain');
    registerDropZone('word-drop-zone', 'word-file-input', wordFiles, 'word-file-list', 'word-file-count', 'word');

    async function handleSelectedFiles(filesList, filesArray, listId, countId, type) {
        if (!filesList || filesList.length === 0) return;

        for (let file of filesList) {
            if (filesArray.some(f => f.file.name === file.name && f.file.size === file.size)) {
                logMsg(`[SYSTEM] 警告：檔案已在列表中，跳過重複檔案: ${file.name}`);
                continue;
            }

            const stem = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
            const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

            const item = {
                file: file,
                id: Math.random().toString(36).substring(2, 9),
                startPage: 1,
                endPage: null,
                maxPages: 1,
                outputName: stem,
                isPdf: ext === '.pdf'
            };

            filesArray.push(item);
            logMsg(`[SYSTEM] 已選取檔案: ${file.name} (${(file.size / (1024 * 1024)).toFixed(2)} MB)`);

            if (item.isPdf) {
                try {
                    const pages = await getPdfPageCount(file);
                    item.maxPages = pages;
                    item.endPage = pages;
                    renderFileList(filesArray, listId, countId, type);
                } catch (err) {
                    logMsg(`[ERROR] 讀取 PDF 頁數失敗 (${file.name}): ${err.message}`, 'error');
                }
            }
        }

        renderFileList(filesArray, listId, countId, type);
    }

    function renderFileList(filesArray, listId, countId, type) {
        const listContainer = document.getElementById(listId);
        const countSpan = document.getElementById(countId);

        countSpan.textContent = filesArray.length;

        if (filesArray.length === 0) {
            listContainer.innerHTML = '<div class="empty-list-text">尚未選取任何檔案...</div>';
            return;
        }

        listContainer.innerHTML = '';

        filesArray.forEach(item => {
            const fileRow = document.createElement('div');
            fileRow.className = 'file-item';
            
            const iconClass = item.isPdf ? 'fa-file-pdf' : (type === 'word' ? 'fa-file-word' : 'fa-file-image');
            const iconColor = item.isPdf ? '#ff5555' : (type === 'word' ? '#2196f3' : '#ffb300');
            const nameLabel = item.file.name;

            let pageSelectorHtml = '';
            let outputNameHtml = '';

            if (item.isPdf && type !== 'word') {
                pageSelectorHtml = `
                    <div class="page-range-input">
                        <span>頁數:</span>
                        <input type="number" value="${item.startPage}" min="1" max="${item.maxPages}" data-field="startPage" data-id="${item.id}">
                        <span>-</span>
                        <input type="number" value="${item.endPage || item.maxPages}" min="1" max="${item.maxPages}" data-field="endPage" placeholder="${item.maxPages}" data-id="${item.id}">
                    </div>
                `;
            }

            if (type !== 'word') {
                outputNameHtml = `
                    <div class="custom-output-name">
                        <span>輸出檔名:</span>
                        <input type="text" value="${item.outputName}" data-field="outputName" data-id="${item.id}">
                        <span>.docx</span>
                    </div>
                `;
            }

            fileRow.innerHTML = `
                <div class="file-item-header">
                    <div class="file-name-info" title="${item.file.name}">
                        <i class="fa-solid ${iconClass}" style="color: ${iconColor};"></i>
                        <span>${nameLabel}</span>
                    </div>
                    <button class="small-btn btn-danger remove-file-btn" data-id="${item.id}" ${isProcessing ? 'disabled' : ''}>
                        <i class="fa-solid fa-xmark"></i>
                    </button>
                </div>
                ${(pageSelectorHtml !== '' || outputNameHtml !== '') ? `
                    <div class="file-item-controls">
                        ${pageSelectorHtml}
                        ${outputNameHtml}
                    </div>
                ` : ''}
            `;

            // Inputs event bindings
            fileRow.querySelectorAll('input').forEach(input => {
                input.addEventListener('change', (e) => {
                    const field = e.target.dataset.field;
                    const id = e.target.dataset.id;
                    const targetItem = filesArray.find(f => f.id === id);
                    if (targetItem) {
                        if (field === 'startPage') {
                            targetItem.startPage = parseInt(e.target.value) || 1;
                        } else if (field === 'endPage') {
                            targetItem.endPage = parseInt(e.target.value) || targetItem.maxPages;
                        } else if (field === 'outputName') {
                            targetItem.outputName = e.target.value.trim();
                        }
                    }
                });
            });

            // Delete item button
            fileRow.querySelector('.remove-file-btn').addEventListener('click', (e) => {
                const id = e.currentTarget.dataset.id;
                const idx = filesArray.findIndex(f => f.id === id);
                if (idx !== -1) {
                    logMsg(`[SYSTEM] 已移除選取檔案: ${filesArray[idx].file.name}`);
                    filesArray.splice(idx, 1);
                    renderFileList(filesArray, listId, countId, type);
                }
            });

            listContainer.appendChild(fileRow);
        });
    }

    document.getElementById('ocr-clear-files-btn').addEventListener('click', () => {
        ocrFiles.length = 0;
        renderFileList(ocrFiles, 'ocr-file-list', 'ocr-file-count', 'ocr');
        logMsg('[SYSTEM] 已清空 OCR 待處理檔案列表。');
    });
    document.getElementById('explain-clear-files-btn').addEventListener('click', () => {
        explainFiles.length = 0;
        renderFileList(explainFiles, 'explain-file-list', 'explain-file-count', 'explain');
        logMsg('[SYSTEM] 已清空詳解待處理檔案列表。');
    });
    document.getElementById('word-clear-files-btn').addEventListener('click', () => {
        wordFiles.length = 0;
        renderFileList(wordFiles, 'word-file-list', 'word-file-count', 'word');
        logMsg('[SYSTEM] 已清空 Word 優化列表。');
    });

    /* ==========================================
       5. Prompt Tuning & Rule Modals
       ========================================== */

    function getRule(type) {
        if (type === 'ocr') {
            return localStorage.getItem('ocr-rule') || DEFAULT_OCR_RULE;
        } else {
            return localStorage.getItem('explain-rule') || DEFAULT_EXPLAIN_RULE;
        }
    }

    function saveRule(type, content) {
        if (type === 'ocr') {
            localStorage.setItem('ocr-rule', content);
        } else {
            localStorage.setItem('explain-rule', content);
        }
    }

    document.getElementById('ocr-edit-rule-btn').addEventListener('click', () => {
        activeRuleType = 'ocr';
        ruleModalTitle.innerHTML = '📝 編輯 OCR 規則 (Prompt)';
        ruleModalTextarea.value = getRule('ocr');
        ruleModal.classList.add('show');
    });

    document.getElementById('explain-edit-rule-btn').addEventListener('click', () => {
        activeRuleType = 'explain';
        ruleModalTitle.innerHTML = '💡 編輯詳解規則 (Prompt)';
        ruleModalTextarea.value = getRule('explain');
        ruleModal.classList.add('show');
    });

    ruleModalClose.addEventListener('click', () => ruleModal.classList.remove('show'));
    
    ruleModalSaveBtn.addEventListener('click', () => {
        saveRule(activeRuleType, ruleModalTextarea.value);
        ruleModal.classList.remove('show');
        showToast('規則已儲存成功！', 'success');
        logMsg(`[SYSTEM] 已更新 ${activeRuleType === 'ocr' ? 'OCR' : '詳解'} Prompt 規則設定。`);
    });

    ruleModalResetBtn.addEventListener('click', () => {
        if (confirm('確定要還原成預設規則嗎？這將會覆寫您的修改。')) {
            const defaultVal = activeRuleType === 'ocr' ? DEFAULT_OCR_RULE : DEFAULT_EXPLAIN_RULE;
            ruleModalTextarea.value = defaultVal;
            showToast('已還原為預設規則', 'info');
        }
    });

    ruleModalAiBtn.addEventListener('click', () => {
        refineBadOutput.value = '';
        refineGoodOutput.value = '';
        refineResultOutput.value = '';
        aiRefineModal.classList.add('show');
    });

    aiRefineModalClose.addEventListener('click', () => aiRefineModal.classList.remove('show'));

    refineGenerateBtn.addEventListener('click', async () => {
        const badText = refineBadOutput.value.trim();
        const goodText = refineGoodOutput.value.trim();
        const oldRule = ruleModalTextarea.value;

        if (!badText || !goodText) {
            showToast('請填寫錯誤輸出與期望輸出！', 'warning');
            return;
        }

        refineGenerateBtn.disabled = true;
        refineResultOutput.value = '🧠 AI 正在思考中，精煉優化 Prompt 提示詞，請稍候...';

        const promptTuningMessage = `你是一個「規則更新器」。
你的任務：
1. 根據 AI 錯誤輸出 upper
2. 根據使用者期望的正確輸出 lower
3. 根據舊規則 rule_text
輸出「更新後的新規則」。
要求：
- 只能輸出新規則本身
- 不要解釋
- 不要加入其它內容
------------------------------------
【AI 錯誤輸出 upper】
${badText}
------------------------------------
【使用者期望輸出 lower】
${goodText}
------------------------------------
【舊規則 rule_text】
${oldRule}
------------------------------------
請輸出：【更新後的新規則】`;

        try {
            const responseText = await callAIApi(promptTuningMessage, 'text');
            refineResultOutput.value = responseText;
        } catch (err) {
            refineResultOutput.value = `❌ AI 優化失敗: ${err.message}`;
        } finally {
            refineGenerateBtn.disabled = false;
        }
    });

    refineApplyBtn.addEventListener('click', () => {
        const newRule = refineResultOutput.value.trim();
        if (newRule && !newRule.includes('AI 正在思考') && !newRule.includes('❌')) {
            ruleModalTextarea.value = newRule;
            aiRefineModal.classList.remove('show');
            showToast('新規則已成功套用！', 'success');
        } else {
            showToast('無有效的新規則可套用！', 'error');
        }
    });

    /* ==========================================
       6. Copy Buttons
       ========================================== */

    document.getElementById('ocr-copy-output-btn').addEventListener('click', () => {
        const area = document.getElementById('ocr-output-area');
        area.select();
        document.execCommand('copy');
        showToast('OCR 辨識內容已複製！', 'success');
    });

    document.getElementById('explain-copy-output-btn').addEventListener('click', () => {
        const area = document.getElementById('explain-output-area');
        area.select();
        document.execCommand('copy');
        showToast('詳解內容已複製！', 'success');
    });

    /* ==========================================
       7. Execution Loops (OCR & Explain)
       ========================================== */

    async function executeOcrProcess(type) {
        const filesArray = type === 'ocr' ? ocrFiles : explainFiles;
        const outputArea = document.getElementById(`${type}-output-area`);
        const startBtn = document.getElementById(`${type}-start-btn`);
        const stopBtn = document.getElementById(`${type}-stop-btn`);
        const globalStart = parseInt(document.getElementById(`${type}-global-start`).value) || null;
        const globalEnd = parseInt(document.getElementById(`${type}-global-end`).value) || null;

        if (filesArray.length === 0) {
            showToast('請先選取待處理檔案！', 'warning');
            return;
        }

        isProcessing = true;
        stopRequested = false;
        startBtn.disabled = true;
        stopBtn.disabled = false;
        stopBtn.classList.remove('hide');
        progressContainer.classList.remove('hide');

        outputArea.value = '';
        
        let processedFilesCount = 0;
        const totalFiles = filesArray.length;

        renderFileList(filesArray, `${type}-file-list`, `${type}-file-count`, type);
        logMsg(`[SYSTEM] 🚀 開始處理 ${totalFiles} 個任務 (${type === 'ocr' ? 'OCR 辨識' : '詳解生成'})...`);

        try {
            const rulePrompt = getRule(type);

            for (let i = 0; i < totalFiles; i++) {
                if (stopRequested) break;

                const item = filesArray[i];
                logMsg(`📄 [${i+1}/${totalFiles}] 正在處理: ${item.file.name}`);

                const isPdf = item.isPdf;
                const outputName = item.outputName || 'output';

                let fileTexts = '';
                let txtTempContent = '';

                if (isPdf) {
                    const startPage = globalStart || item.startPage || 1;
                    const endPage = globalEnd || item.endPage || item.maxPages;

                    logMsg(`📄 PDF 共 ${item.maxPages} 頁，設定轉換範圍: ${startPage} - ${endPage} 頁`);

                    for (let page = startPage; page <= endPage; page++) {
                        if (stopRequested) break;

                        logMsg(`   -> 正在渲染並轉換第 ${page}/${endPage} 頁...`);
                        const percent = Math.floor(((processedFilesCount / totalFiles) + ((page - startPage) / (endPage - startPage + 1)) / totalFiles) * 100);
                        updateProgress(`📄 處理中 (${page}/${endPage} 頁): ${item.file.name}`, percent);

                        try {
                            const base64Png = await convertPdfPageToPngBase64(item.file, page);
                            const marker = `##### Page ${page} #####`;
                            const pageText = await callAIApi(rulePrompt, 'image', base64Png);
                            
                            if (pageText) {
                                fileTexts += `${pageText}\n\n`;
                                txtTempContent += `${marker}\n${pageText}\n\n`;
                            }
                        } catch (err) {
                            logMsg(`❌ 第 ${page} 頁轉換失敗: ${err.message}`, 'error');
                        }
                    }
                } else {
                    updateProgress(`🖼️ 處理圖片: ${item.file.name}`, Math.floor((processedFilesCount / totalFiles) * 100));
                    try {
                        const base64Png = await readImageAsBase64(item.file);
                        const ocrResult = await callAIApi(rulePrompt, 'image', base64Png);
                        if (ocrResult) {
                            fileTexts += `${ocrResult}\n\n`;
                            txtTempContent += `--- ${item.file.name} ---\n${ocrResult}\n\n`;
                        }
                    } catch (err) {
                        logMsg(`❌ 圖片轉換失敗: ${err.message}`, 'error');
                    }
                }

                if (stopRequested) break;

                outputArea.value += `--- ${item.file.name} ---\n${fileTexts}\n\n`;
                outputArea.scrollTop = outputArea.scrollHeight;

                if (fileTexts.trim() !== '') {
                    logMsg(`💾 生成 Word 檔案: ${outputName}.docx...`);
                    const wordCleanText = fileTexts.replace(/##### Page \d+ #####\n?/g, '');
                    await buildWordDocument(wordCleanText, outputName);
                    triggerTxtDownload(txtTempContent, `${outputName}.txt`);
                }

                processedFilesCount++;
                updateProgress(`✅ 已處理完成: ${item.file.name}`, Math.floor((processedFilesCount / totalFiles) * 100));
            }

            if (stopRequested) {
                logMsg('[SYSTEM] 🛑 任務已被使用者強制中斷！', 'error');
                showToast('任務已中斷', 'warning');
            } else {
                logMsg('✅ ✅ ✅ 所有檔案處理完畢！已全部導出為 Word 下載。');
                showToast('轉換成功完成！🌻', 'success');
            }

        } catch (err) {
            logMsg(`❌ 發生致命錯誤: ${err.message}`, 'error');
            showToast(err.message, 'error');
        } finally {
            isProcessing = false;
            startBtn.disabled = false;
            stopBtn.disabled = true;
            stopBtn.classList.add('hide');
            
            setTimeout(() => {
                progressContainer.classList.add('hide');
            }, 3000);

            renderFileList(filesArray, `${type}-file-list`, `${type}-file-count`, type);
        }
    }

    function updateProgress(taskName, percent) {
        progressTaskName.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> ${taskName}`;
        progressPercentage.textContent = `${percent}%`;
        progressBarFill.style.width = `${percent}%`;
    }

    document.getElementById('ocr-start-btn').addEventListener('click', () => executeOcrProcess('ocr'));
    document.getElementById('explain-start-btn').addEventListener('click', () => executeOcrProcess('explain'));

    document.getElementById('ocr-stop-btn').addEventListener('click', () => requestStop('ocr'));
    document.getElementById('explain-stop-btn').addEventListener('click', () => requestStop('explain'));

    function requestStop(type) {
        stopRequested = true;
        logMsg('[SYSTEM] 🛑 正在請求停止，請稍候於當前頁面完成後結束...');
        document.getElementById(`${type}-stop-btn`).disabled = true;
        document.getElementById(`${type}-stop-btn`).textContent = '🛑 停止中...';
    }

    /* ==========================================
       8. Word Refinement Handler
       ========================================== */

    const wordStartBtn = document.getElementById('word-start-btn');

    wordStartBtn.addEventListener('click', async () => {
        if (wordFiles.length === 0) {
            showToast('請先選取待排版優化的 Word 檔案！', 'warning');
            return;
        }

        wordStartBtn.disabled = true;
        wordStartBtn.textContent = '⏳ 優化處理中...';

        const options = {
            convertFont: document.getElementById('word-convert-font').checked,
            convertItalic: document.getElementById('word-convert-italic').checked,
            minusToHyphen: document.getElementById('word-minus-to-hyphen').checked,
            hyphenToMinus: document.getElementById('word-hyphen-to-minus').checked,
            convertSuper: document.getElementById('word-convert-super').checked,
            convertSub: document.getElementById('word-convert-sub').checked
        };

        logMsg(`\n[SYSTEM] 🚀 開始對 ${wordFiles.length} 個 Word 文件進行排版優化...`);

        try {
            for (let item of wordFiles) {
                logMsg(`📄 優化檔案: ${item.file.name}`);
                await optimizeWordFile(item.file, options, logMsg);
            }
            showToast('Word 排版優化完成！', 'success');
            logMsg('[SYSTEM] 所有 Word 檔案排版優化程序執行完畢。🌻');
        } catch (err) {
            logMsg(`❌ Word 優化失敗: ${err.message}`, 'error');
            showToast(err.message, 'error');
        } finally {
            wordStartBtn.disabled = false;
            wordStartBtn.textContent = '🚀 開始執行優化';
        }
    });
});
