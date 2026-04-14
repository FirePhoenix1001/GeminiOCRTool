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

from gemini import GeminiProcessor
from gpt import GPTProcessor

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
        self.stop_requested = False
        self.processor = None
        
        # 綁定按鈕
        self.app.start_btn.configure(command=lambda: self.run_thread(mode="ocr"))
        self.app.stop_btn.configure(command=self.stop_process)
        
        if hasattr(self.app, 'explain_start_btn'):
            self.app.explain_start_btn.configure(command=lambda: self.run_thread(mode="explain"))
        if hasattr(self.app, 'explain_stop_btn'):
            self.app.explain_stop_btn.configure(command=self.stop_process)

    def log(self, msg, mode="ocr"):
        if mode == "ocr":
            self.app.log(msg)
        else:
            self.app.explain_log(msg)

    def stop_process(self):
        """ 發送停止請求 """
        self.stop_requested = True
        self.log("\n🛑 正在請求停止，請稍候於當前頁面完成後結束...", mode="ocr")
        self.log("\n🛑 正在請求停止，請稍候於當前頁面完成後結束...", mode="explain")

    def run_thread(self, mode="ocr"):
        files = self.app.get_selected_files() if mode == "ocr" else self.app.get_explain_selected_files()
        
        if not files:
            msg = "⚠️ 請選取檔案！"
            if mode == "ocr": self.app.log(msg)
            else: self.app.explain_log(msg)
            return

        self.stop_requested = False
        
        # UI 狀態轉換
        if mode == "ocr":
            self.app.start_btn.configure(state="disabled", text="⏳ 處理中...")
            self.app.stop_btn.configure(state="normal")
            self.app.set_all_items_processing_mode(True, mode="ocr")
        else:
            self.app.explain_start_btn.configure(state="disabled", text="⏳ 處理中...")
            self.app.explain_stop_btn.configure(state="normal")
            self.app.set_all_items_processing_mode(True, mode="explain")

        
        threading.Thread(target=self.run_process, args=(mode,), daemon=True).start()

    def run_process(self, mode="ocr"):
        try:
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

            self.log(f"", mode=mode)
            try:
                if "gpt" in model_name.lower():
                    self.processor = GPTProcessor(api_keys_text, model_name, rule_text, log_callback=lambda m: self.log(m, mode=mode))
                    identify_func = self.processor.gpt_identify
                else:
                    self.processor = GeminiProcessor(api_keys_text, model_name, rule_text, log_callback=lambda m: self.log(m, mode=mode))
                    identify_func = self.processor.gemini_identify
            except Exception as e:
                self.log(f"❌ 初始化失敗: {e}", mode=mode)
                return

            total_files = len(files)
            self.log(f"", mode=mode)

            for idx, file_path in enumerate(files):
                if self.stop_requested: break
                
                filename = os.path.basename(file_path)
                self.log(f"正在處理 ({idx+1}/{total_files}): {filename}", mode=mode)

                ext = os.path.splitext(filename)[1].lower()
                if ext == ".pdf":
                    f_set = file_settings.get(file_path, {"start": 1, "end": 0, "output_name": None})
                    self.process_pdf(file_path, filename, f_set["start"], f_set["end"], identify_func, mode=mode, output_name=f_set.get("output_name"))
                elif ext in [".png", ".jpg", ".jpeg"]:
                    f_set = file_settings.get(file_path, {"output_name": None})
                    self.process_image(file_path, filename, identify_func, mode=mode, output_name=f_set.get("output_name"))
                else:
                    self.log(f"⚠️ 跳過不支援格式: {filename}", mode=mode)

            if self.stop_requested:
                self.log("\n🛑 任務已由使用者手動中斷。", mode=mode)
            else:
                self.log("\n✅ ✅ ✅ 所有任務執行完畢！", mode=mode)

        except Exception as e:
            self.log(f"\n❌ 發生錯誤: {e}", mode=mode)
            traceback.print_exc()
        finally:
            self.processor = None
            if mode == "ocr":
                self.app.start_btn.configure(state="normal", text="開始進行轉換")
                self.app.stop_btn.configure(state="disabled")
                self.app.set_all_items_processing_mode(False, mode="ocr")
            else:
                self.app.explain_start_btn.configure(state="normal", text="開始進行轉換")
                self.app.explain_stop_btn.configure(state="disabled")
                self.app.set_all_items_processing_mode(False, mode="explain")


    def process_image(self, file_path, filename, identify_func, mode="ocr", output_name=None):
        try:
            # 如果 output_name 沒給，預設就是 base_name
            if not output_name:
                output_name = base_name
            else:
                # 去除可能帶有的 .doc 或 .docx 後綴
                if output_name.lower().endswith(".docx"): output_name = output_name[:-5]
                if output_name.lower().endswith(".doc"): output_name = output_name[:-4]

            # 暫存檔名改為以自定義名稱為主
            txt_name = f"{output_name}.txt"
            
            # 檢查是否已經處理過且有內容 (雖然單張圖片通常不接續，但為了邏輯統一)
            existing_text = ""
            if os.path.exists(txt_name):
                with open(txt_name, "r", encoding="utf-8") as f:
                    existing_text = f.read().strip()
            
            if existing_text:
                self.log(f"ℹ️ {filename} 已有現存辨識內容，將直接使用現有內容。", mode=mode)
                result_text = existing_text
            else:
                result_text = identify_func(file_path)
                if result_text:
                    with open(txt_name, "w", encoding="utf-8") as f:
                        f.write(result_text)
            
            if result_text:
                if mode == "ocr":
                    self.app.append_output(f"--- {filename} ---\n{result_text}\n")
                else:
                    self.app.append_explain_output(f"--- {filename} ---\n{result_text}\n")
                
                # 圖片也支援轉 Word (如果使用者需要)
                if inputWord:
                    # 重新從文字文件讀取內容 (以文字文件為準)
                    with open(txt_name, "r", encoding="utf-8") as f:
                        final_text = f.read()
                    self.log(f"✅ 寫入 Word: {output_name}.docx", mode=mode)
                    inputWord(final_text, output_name)
                    
                self.log(f"✅ {filename} 處理完成", mode=mode)
                
                # 辨識成功完成後，刪除暫存的 .txt (除非意外暫停或停止，目前是非 stop 狀態)
                if not self.stop_requested and os.path.exists(txt_name):
                    try: os.remove(txt_name)
                    except: pass
            else:
                self.log(f"⚠️ {filename} 結果為空", mode=mode)
        except Exception as e:
            self.log(f"❌ {filename} 失敗: {e}", mode=mode)

    def process_pdf(self, file_path, filename, start_page, end_page, identify_func, mode="ocr", output_name=None):
        if not pdf_to_picture:
            self.log("❌ 找不到 pdfToPicture 模組", mode=mode)
            return

        self.log(f"", mode=mode)
        try:
            if os.path.exists("picture"): shutil.rmtree("picture")
            pdf_to_picture(file_path, start_page=start_page, end_page=end_page)
            
            if not os.path.exists("picture"):
                self.log("❌ PDF 轉圖失敗", mode=mode)
                return

            def extract_page_num_local(fname):
                match = re.search(r'page_(\d+)', fname)
                return int(match.group(1)) if match else 0

            img_files = sorted(os.listdir('picture'), key=extract_page_num_local)
            self.log(f"📄 共 {len(img_files)} 頁，開始處理...", mode=mode)

            pdf_basename = os.path.splitext(filename)[0]
            # 如果 output_name 沒給，預設就是 pdf_basename
            if not output_name:
                output_name = pdf_basename
            else:
                if output_name.lower().endswith(".docx"): output_name = output_name[:-5]
                if output_name.lower().endswith(".doc"): output_name = output_name[:-4]

            # 暫存檔名改為以自定義名稱為主
            txt_name = f"{output_name}.txt"
            
            # 讀取現有內容
            existing_content = ""
            if os.path.exists(txt_name):
                with open(txt_name, "r", encoding="utf-8") as f:
                    existing_content = f.read()
            
            for img in img_files:
                if self.stop_requested: break
                
                # 提取當前頁碼
                page_n = extract_page_num_local(img)
                marker = f"##### Page {page_n} #####"
                
                # 檢查是否已經在文字檔案中
                if marker in existing_content:
                    self.log(f"   -> 跳過已存在頁面: {img}", mode=mode)
                    continue

                img_path = os.path.join("picture", img)
                self.log(f"   -> 處理頁面: {img}", mode=mode)
                page_text = identify_func(img_path)
                
                if page_text:
                    with open(txt_name, "a", encoding="utf-8") as f:
                        f.write(f"{marker}\n{page_text}\n\n")
            
            # 最後以文字文件的內容為主，生成 Word
            if not self.stop_requested:
                if os.path.exists(txt_name):
                    with open(txt_name, "r", encoding="utf-8") as f:
                        final_raw_content = f.read()
                    
                    # 過濾掉頁碼標記，保持 Word 內容乾淨
                    final_content_for_word = re.sub(r'##### Page \d+ #####\n?', '', final_raw_content)
                        
                    if inputWord:
                        self.log(f"✅ 寫入 Word: {output_name}.docx", mode=mode)
                        inputWord(final_content_for_word, output_name)

                # 辨識成功完成後，刪除暫存的 .txt (除非意外暫停或停止)
                if not self.stop_requested and os.path.exists(txt_name):
                    try: os.remove(txt_name)
                    except: pass
            
            if os.path.exists("picture"): shutil.rmtree("picture")
            # 不再刪除 txt_name，以便接續與供使用者修改後查看
            # if os.path.exists(txt_name): os.remove(txt_name)

        except Exception as e:
            self.log(f"❌ PDF 處理失敗: {e}", mode=mode)

if __name__ == "__main__":
    import customtkinter as ctk
    from gui import GeminiOCRApp
    root = ctk.CTk()
    app = GeminiOCRApp(root, default_rule_text=RULE, default_explain_rule_text=RULE_EXPLAIN)
    controller = OCRController(app) 
    root.mainloop()

# ==========================================
# 程式進入點 (CLI 模式 - 開發者除錯)
# ==========================================
# 想要跳過 GUI 時，請手動將上方改為 "__main1__"
elif __name__ == "__main__":
    from gemini import setup_gemini, gemini_identify
    from toWord import changeWord, inputWord
    print("--- 進入 CLI / 開發者模式 ---")
    
    api_key_dict = {} 
    model = "gemini-1.5-flash"
    
    try:
        try:
            setup_gemini(api_key_dict, model, RULE)
        except ValueError:
            print("注意：未提供 Key，請確保環境變數已設定。")

        def extract_page_number(filename):
            match = re.search(r'page_(\d+)', filename)
            return int(match.group(1)) if match else 0

        pdf_files = [f for f in os.listdir('.') if f.lower().endswith('.pdf')]
        png_files = [f for f in os.listdir('.') if f.lower().endswith('.png')]
        word_files = [f for f in os.listdir('.') if f.lower().endswith(('.docx', '.doc'))]

        if word_files:
            for file in word_files:
                if changeWord: changeWord(file)

        if png_files:
            for file in png_files:
                text = gemini_identify(str(file))
                with open(f"OUTPUT.txt", "a", encoding="utf-8") as f:
                    f.write(str(text) + "\n\n")
                print(f"完成 {file}")
            
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

    except Exception as e:
        print(f"❌ 發生錯誤: {e}")
        traceback.print_exc()

    input("按下Enter鍵結束程式")