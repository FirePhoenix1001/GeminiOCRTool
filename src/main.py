import os
import sys
import threading
import shutil
import re
import time
import traceback

# 嘗試引用周邊模組 (防呆機制)
try:
    from pdfToPicture import pdf_to_picture
except ImportError:
    pdf_to_picture = None

try:
    from toWord import inputWord, changeWord
except ImportError:
    inputWord = None
    changeWord = None

from gemini import gemini_identify, setup_gemini

# ==========================================
# 預設規則 (Prompt)
# ==========================================
RULE = r"""<role>
你是一位追求「100% 忠實還原」的 OCR 數學數位化專家，擅長以「純文字為主、LaTeX 為輔」的方式排版。
</role>

<rules>
1. **LaTeX 使用的絕對禁令（環境繼承原則）：**
   - **唯一例外：** ** 只有在出現「分數 \frac{}{}」、「根號 \sqrt{}{}」及「**聯立方程式（如 \begin{cases}）**」時，才允許使用 $...$。
   - **環境繼承：** 一旦進入 $ 內部，所有內容（含運算符、上下標、變數）必須統一採用 LaTeX 寫法（如 \times, n_{1}）。

2. **純文字優先與簡化標記（最高準則）：**
   - **一般變數：** 所有的變數、數字、希臘字母、邏輯符號、比較符號，只要在 $ 之外，一律採純文字，嚴禁包進 $ 內。
   - **上下標處理：** 在 $ 之外時，直接使用 `_{}` 表示下標，`^{}` 表示上標（例如：L_{1}, x^{2}, (y−2)^{2}）。
   - **範例：** 使用 x≈y±z 而不是 $x$≈$y$±$z$。

3. **符號清單（限純文字顯示）：**
   - **關係：** <, >, =, ≤, ≥, ≠, ≈, ±。
   - **邏輯：** ∵, ∴, ⇒, ⇔, ∈。
   - **運算：** ×, ·, π, …, ∠, $\overline{AB}$ (上劃線仍可用 $)。
   - **希臘/序號：** θ, α, β, Γ, ①, ②。
   - **函數：** log, sin, cos, tan。

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
  錯誤：$x$≈$y$±$z$ 或 $x$∈$A$ ⇒ $x$≠$①$

- **幾何與邏輯：**
  正確：∵ ∠α=∠β ∴ △ABC 為等腰三角形。
  正確：$\overline{AB}$=$\overline{AC}$

- **必須使用 LaTeX 的情境：**
  正確：$\frac{1}{2}$+$\frac{\sqrt{3}}{2}$ = $\frac{1+\sqrt{3}}{2}$
  注意：只有分數 and 根號區塊才使用 $，中間的加號與等號若不在 $ 內，請視情況而定。若進入 $ 內部，則遵守規則 1.2（採 LaTeX 寫法）。
  正確：$\frac{\pi}{2}$+$\frac{\pi}{2}$ = π
  正確：(A) $\begin{cases} a+2\theta\beta>0 \\ 7a-24b\theta\beta+22>0 \end{cases}$
  （注意：選項 (A) 在 $ 外保持純文字；聯立方程內部符號如 > 採 LaTeX 寫法）

- **函數與運算：**
  正確：log 2 + sin θ = 1
  正確：π ≈ 3.14
  錯誤：$log$ 2 + $sin$ $θ$ = 1 或 $π$ ≈ 3.14

- **一般上下標（使用純文字標記）：**
  正確：L_{1}: 2x−y+a=0
  正確：(x+3)^{2}+(y−2)^{2}=60
  錯誤：$L_{1}$ 或 $(x+3)^2$
  錯誤：L_1 或 $(x+3)^2$

- **分數/根號環境（使用 LaTeX 標記）：**
  正確：$\frac{\pi n_{1}}{\pi}$ = n_{1}
  （解釋：因為在分數內，所以 n_{1} 必須跟著進入 $ 環境並使用 LaTeX 語法；但等號後的結果回歸純文字標記。）
</examples>

<task>
根據上述「純文字標記優先」規則，將提供圖片內容轉錄為文字。
你的核心任務是將圖片中的數學內容「完全相符」地轉錄為文字，嚴禁任何擅自的簡化、省略 or 美化。
</task>"""

RULE_EXPLAIN = r"""<role>
你是一位極其專業且耐心的「數學詳解專家」，擅長以「純文字為主、LaTeX 為輔」的方式，為數學題目提供步驟詳盡、邏輯嚴密的解答。
你的目標是：不僅給出正確答案，更要讓學生看懂解題的每一個轉折。
</role>

<rules>
1. **LaTeX 使用的絕對禁令（環境繼承原則）：**
   - **唯一例外：** ** 只有在出現「分數 \frac{}{}」、「根號 \sqrt{}{}」及「**聯立方程式（如 \begin{cases}）**」時，才允許使用 $...$。
   - **環境繼承：** 一旦進入 $ 內部，所有內容（含運算符、上下標、變數）必須統一採用 LaTeX 寫法（如 \times, n_{1}）。

2. **純文字優先與簡化標記（最高準則）：**
   - **一般變數：** 所有的變數、數字、希臘字母、邏輯符號、比較符號，只要在 $ 之外，一律採純文字，嚴禁包進 $ 內。
   - **上下標處理：** 在 $ 之外時，直接使用 `_{}` 表示下標，`^{}` 表示上標（例如：L_{1}, x^{2}, (y−2)^{2}）。
   - **範例：** 使用 x≈y±z 而不是 $x$≈$y$±$z$。

3. **符號清單（限純文字顯示）：**
   - **關係：** <, >, =, ≤, ≥, ≠, ≈, ±。
   - **邏輯：** ∵, ∴, ⇒, ⇔, ∈。
   - **運算：** ×, ·, π, …, ∠, $\overline{AB}$ (上劃線仍可用 $)。
   - **希臘/序號：** θ, α, β, Γ, ①, ②。
   - **函數：** log, sin, cos, tan。

4. **解題結構：**
   - **題目：** 先簡短摘要題目內容。
   - **詳解：** 分步驟說明解法，推導邏輯需流暢且明確。
   - **答案：** 在最後明確標註最終答案。
   - **禁令：** 不使用任何 Markdown 格式語法。
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
</task>"""


# ==========================================
# 控制器 (Controller) - 負責 GUI 邏輯與多執行緒
# ==========================================
class OCRController:
    def __init__(self, app):
        self.app = app
        # 【關鍵步驟】覆寫 GUI 的按鈕指令，指向這裡的 run_thread
        self.app.start_btn.configure(command=lambda: self.run_thread(mode="ocr"))
        
        # 如果有詳解按鈕，也一併綁定
        if hasattr(self.app, 'explain_start_btn'):
            self.app.explain_start_btn.configure(command=lambda: self.run_thread(mode="explain"))

    def log(self, msg, mode="ocr"):
        """ 轉發訊息給 GUI 的 Log 視窗 """
        if mode == "ocr":
            self.app.log(msg)
        else:
            self.app.explain_log(msg)

    def run_thread(self, mode="ocr"):
        """ 啟動多執行緒，避免 GUI 卡死 """
        files = self.app.get_selected_files() if mode == "ocr" else self.app.get_explain_selected_files()
        
        if not files:
            msg = "⚠️ 請先選擇至少一個檔案！"
            if mode == "ocr": self.app.log(msg)
            else: self.app.explain_log(msg)
            return

        # 鎖定按鈕，顯示處理中
        if mode == "ocr":
            self.app.start_btn.configure(state="disabled", text="⏳ 正在處理中...")
        else:
            self.app.explain_start_btn.configure(state="disabled", text="⏳ 正在處理中...")
        
        # 開啟背景執行緒跑 run_process
        threading.Thread(target=self.run_process, args=(mode,), daemon=True).start()

    def run_process(self, mode="ocr"):
        """ 真正執行 OCR 的邏輯 (背景執行) """
        try:
            # 1. 取得 GUI 設定
            if mode == "ocr":
                files = self.app.get_selected_files()
                rule_text = self.app.get_rule_text()
                file_settings = self.app.get_file_settings()
            else:
                files = self.app.get_explain_selected_files()
                rule_text = self.app.get_explain_rule_text()
                file_settings = self.app.get_explain_file_settings()

            api_keys_text = self.app.get_api_key()
            model_name = self.app.get_model()

            # 2. 初始化 Gemini (並傳入 self.log 讓錯誤訊息能回傳 GUI)
            self.log(f"\n--- 🚀 開始初始化 Gemini ({model_name}) ---", mode=mode)
            try:
                setup_gemini(api_keys_text, model_name, rule_text, log_callback=lambda m: self.log(m, mode=mode))
            except Exception as e:
                self.log(f"❌ 初始化失敗: {e}", mode=mode)
                return # 終止

            # 3. 逐一處理檔案
            total_files = len(files)
            self.log(f"--- 📂 準備處理 {total_files} 個檔案 ---\n", mode=mode)

            for idx, file_path in enumerate(files):
                filename = os.path.basename(file_path)
                self.log(f"正在處理 ({idx+1}/{total_files}): {filename}", mode=mode)

                # 判斷副檔名
                ext = os.path.splitext(filename)[1].lower()

                if ext == ".pdf":
                    f_set = file_settings.get(file_path, {"start": 1, "end": 0})
                    self.process_pdf(file_path, filename, f_set["start"], f_set["end"], mode=mode)
                elif ext in [".png", ".jpg", ".jpeg"]:
                    self.process_image(file_path, filename, mode=mode)
                else:
                    self.log(f"⚠️ 跳過不支援的格式: {filename}", mode=mode)

            self.log("\n✅ ✅ ✅ 所有任務執行完畢！", mode=mode)

        except Exception as e:
            self.log(f"\n❌ 發生嚴重錯誤: {e}", mode=mode)
            traceback.print_exc() # 同時印在終端機方便除錯
        
        finally:
            # 無論成功或失敗，最後都要恢復按鈕
            if mode == "ocr":
                self.app.start_btn.configure(state="normal", text="儲存設定並開始執行")
            else:
                self.app.explain_start_btn.configure(state="normal", text="儲存設定並開始執行")

    # --- 個別檔案處理邏輯 ---

    def process_image(self, file_path, filename, mode="ocr"):
        try:
            result_text = gemini_identify(file_path)
            if result_text:
                if mode == "ocr":
                    self.app.append_output(f"--- {filename} ---\n{result_text}\n")
                else:
                    self.app.append_explain_output(f"--- {filename} ---\n{result_text}\n")
                self.log(f"✅ {filename} 處理完成，已輸出至畫面", mode=mode)
            else:
                self.log(f"⚠️ {filename} 辨識結果為空", mode=mode)
        except Exception as e:
            self.log(f"❌ {filename} 失敗: {e}", mode=mode)

    def process_pdf(self, file_path, filename, start_page=1, end_page=0, mode="ocr"):
        if not pdf_to_picture:
            self.log("❌ 找不到 pdfToPicture 模組，跳過 PDF", mode=mode)
            return

        self.log(f"🔄 正在將 {filename} 轉換為圖片...", mode=mode)
        try:
            if os.path.exists("picture"): shutil.rmtree("picture")
            
            pdf_to_picture(file_path, start_page=start_page, end_page=end_page) # 呼叫模組
            
            if not os.path.exists("picture"):
                self.log("❌ PDF 轉圖失敗", mode=mode)
                return

            def extract_page_num(fname):
                match = re.search(r'page_(\d+)', fname)
                return int(match.group(1)) if match else 0

            img_files = sorted(os.listdir('picture'), key=extract_page_num)
            self.log(f"📄 共 {len(img_files)} 頁，開始處理...", mode=mode)

            pdf_basename = os.path.splitext(filename)[0]
            txt_name = f"{pdf_basename}.txt"
            
            # 確保檔案為空
            if os.path.exists(txt_name):
                os.remove(txt_name)
            
            full_text = ""
            for img in img_files:
                img_path = os.path.join("picture", img)
                self.log(f"   -> 處理頁面: {img}", mode=mode)
                page_text = gemini_identify(img_path)
                if page_text:
                    full_text += page_text + "\n\n"
                    # 漸進式即時寫入文字檔
                    with open(txt_name, "a", encoding="utf-8") as f:
                        f.write(page_text + "\n\n")
                    self.log(f"      (已即時儲存至 {txt_name})", mode=mode)
            
            # 增加輸出到 GUI
            if mode == "ocr":
                self.app.append_output(f"--- {filename} ---\n{full_text}\n")
            else:
                self.app.append_explain_output(f"--- {filename} ---\n{full_text}\n")

            # 轉 Word
            if inputWord:
                self.log(f"✅ 寫入 Word: {pdf_basename}.docx", mode=mode)
                inputWord(full_text, pdf_basename)
            
            # 清理
            shutil.rmtree("picture")
            if os.path.exists(txt_name): os.remove(txt_name)

        except Exception as e:
            self.log(f"❌ PDF 處理失敗: {e}", mode=mode)



# ==========================================
# 程式進入點 (GUI 模式 - 打包/一般使用)
# ==========================================


# [Image of MVC architecture diagram]

if __name__ == "__main__":
    import customtkinter as ctk
    from gui import GeminiOCRApp # 引入你的 gui.py
    
    root = ctk.CTk()
    
    # 1. 啟動 GUI (傳入雙規則)
    app = GeminiOCRApp(root, default_rule_text=RULE, default_explain_rule_text=RULE_EXPLAIN)
    
    # 2. 【關鍵】啟動控制器
    # 這裡就是把「功能」注入到 GUI 的地方
    controller = OCRController(app) 
    
    root.mainloop()


# ==========================================
# 程式進入點 (CLI 模式 - 開發者除錯)
# ==========================================
# 想要跳過 GUI 時，請手動將上方改為 "__main1__"
elif __name__ == "__main__":
    print("--- 進入 CLI / 開發者模式 ---")
    
    # 模擬環境
    api_key_dict = {} 
    model = "gemini-3-flash-preview"
    
    try:
        # 1. 初始化
        try:
            setup_gemini(api_key_dict, model, RULE)
        except ValueError:
            print("注意：未提供 Key，請確保環境變數已設定，否則稍後會出錯。")

        # 2. 掃描檔案
        def extract_page_number(filename):
            match = re.search(r'page_(\d+)', filename)
            return int(match.group(1)) if match else 0

        pdf_files = [f for f in os.listdir('.') if f.lower().endswith('.pdf')]
        png_files = [f for f in os.listdir('.') if f.lower().endswith('.png')]
        word_files = [f for f in os.listdir('.') if f.lower().endswith(('.docx', '.doc'))]

        # 3. 執行邏輯 (包在 try-except 區塊中)
        # -------------------------------------------------------------------
        # 這裡就是修改的重點：捕捉 RuntimeError 來避免印出 Traceback
        # -------------------------------------------------------------------
        
        # Word
        if word_files:
            for file in word_files:
                if changeWord: changeWord(file)

        # PNG
        if png_files:
            for file in png_files:
                text = gemini_identify(str(file))
                with open(f"OUTPUT.txt", "a", encoding="utf-8") as f:
                    f.write(str(text) + "\n\n")
                print(f"完成 {file}")
            print("完成 IMG 轉換")
            
        # PDF
        if pdf_files:
            for file in pdf_files:
                pdf_name=os.path.splitext(file)[0]
                if pdf_to_picture:
                    pdf_to_picture(file)
                    print(f"{len(os.listdir('picture'))} 圖片，請稍後")
                    for img_file in sorted(os.listdir('picture'), key=extract_page_number):
                        text = gemini_identify(f'picture/{img_file}')
                        if text:
                            with open(f"{pdf_name}.txt", "a", encoding="utf-8") as f:
                                f.write(str(text) + "\n\n")
                        print(f"完成 {img_file}")

                    with open(f"{pdf_name}.txt", "r", encoding="utf-8") as f: text = f.read()
                    shutil.rmtree("picture")
                    os.remove(f"{pdf_name}.txt")
                    if inputWord: inputWord(text, pdf_name)
                    print(f"✅ 完成 --- {file}")

        if not (pdf_files or png_files or word_files):
            print(f"\n{'-'*30}\n這裡甚麼都沒有請再試一次\n{'-'*30}\n")

    except RuntimeError as e:
        # 當 gemini.py 拋出 "All API Keys exhausted" 時，會進入這裡
        # 我們只印出乾淨的錯誤訊息，而不印 Stack Trace
        print(f"\n{'!'*40}")
        print(f"🛑 程式終止：{e}")
        print(f"{'!'*40}\n")
        sys.exit(1)
        
    except Exception as e:
        # 其他未預期的錯誤，還是印出來方便除錯
        print(f"❌ 發生未預期錯誤: {e}")
        traceback.print_exc()

    input("按下Enter鍵結束程式")