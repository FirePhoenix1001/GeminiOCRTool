import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
import winreg
import threading
from toWord import changeWord

# ==========================================
# 自定義元件區
# ==========================================

class ModelInfoPopup(ctk.CTkToplevel):
    def __init__(self, parent, title, message):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x350")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.focus_set()
        ctk.CTkLabel(self, text=title, font=("微軟正黑體", 18, "bold")).pack(pady=(20, 10))
        text_frame = ctk.CTkFrame(self, fg_color="transparent")
        text_frame.pack(fill="both", expand=True, padx=20, pady=5)
        self.info_text = ctk.CTkTextbox(text_frame, font=("微軟正黑體", 14), wrap="word", fg_color="transparent")
        self.info_text.pack(fill="both", expand=True)
        self.info_text.insert("1.0", message)
        self.info_text.configure(state="disabled")
        self.ok_btn = ctk.CTkButton(self, text="我知道了", width=120, height=40, font=("微軟正黑體", 14, "bold"), command=self.destroy)
        self.ok_btn.pack(side="bottom", pady=20)


class EnvVarWindow(ctk.CTkToplevel):
    def __init__(self, parent, on_save_callback):
        super().__init__(parent)
        self.title("🌐 設定環境變數 API Key")
        self.geometry("500x250")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.focus_set()
        self.on_save_callback = on_save_callback
        ctk.CTkLabel(self, text="設定系統環境變數 (GEMINI_API_KEY)", font=("微軟正黑體", 16, "bold")).pack(pady=(20, 10))
        ctk.CTkLabel(self, text="留空並儲存即可「完全刪除」此變數。", font=("微軟正黑體", 12), text_color="gray").pack(pady=(0, 20))
        input_frame = ctk.CTkFrame(self, fg_color="transparent")
        input_frame.pack(fill="x", padx=20)
        self.is_hidden = True
        current_env = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or ""
        self.key_entry = ctk.CTkEntry(input_frame, width=350, placeholder_text="請貼上 API Key (留空則刪除)", show="•")
        self.key_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.key_entry.insert(0, current_env)
        self.eye_btn = ctk.CTkButton(input_frame, text="👁", width=40, fg_color="#555555", hover_color="#777777", command=self.toggle_visibility)
        self.eye_btn.pack(side="left", padx=5)
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=30)
        ctk.CTkButton(btn_frame, text="確認設定", width=120, command=self.save_env).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame, text="取消", width=80, fg_color="gray", command=self.destroy).pack(side="right", padx=5)

    def toggle_visibility(self):
        self.is_hidden = not self.is_hidden
        self.key_entry.configure(show="•" if self.is_hidden else "")
        self.eye_btn.configure(fg_color="#555555" if self.is_hidden else "#2CC985")

    def save_env(self):
        key = self.key_entry.get().strip()
        if not key:
            if "GEMINI_API_KEY" in os.environ:
                del os.environ["GEMINI_API_KEY"]
            if os.path.exists(".env"):
                try: os.remove(".env")
                except: pass
            try:
                subprocess.run(r'REG DELETE "HKCU\Environment" /V GEMINI_API_KEY /F', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except: pass
            messagebox.showinfo("提示", "已清除環境變數。")
        else:
            os.environ["GEMINI_API_KEY"] = key
            try:
                subprocess.run(f'setx GEMINI_API_KEY "{key}"', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except: pass
        self.on_save_callback()
        self.destroy()

class APIKeyRow(ctk.CTkFrame):
    def __init__(self, parent, name_val="", key_val="", delete_command=None):
        super().__init__(parent, fg_color="transparent")
        self.pack(fill="x", pady=2)
        self.name_entry = ctk.CTkEntry(self, width=110, placeholder_text="名稱 (選填)")
        self.name_entry.pack(side="left", padx=(0, 5))
        if name_val: self.name_entry.insert(0, name_val)
        self.is_hidden = True
        self.key_entry = ctk.CTkEntry(self, width=260, placeholder_text="API Key", show="•")
        self.key_entry.pack(side="left", fill="x", expand=True, padx=5)
        if key_val: self.key_entry.insert(0, key_val)
        self.eye_btn = ctk.CTkButton(self, text="👁", width=30, fg_color="#555555", command=self.toggle_visibility)
        self.eye_btn.pack(side="left", padx=2)
        self.del_btn = ctk.CTkButton(self, text="❌", width=30, fg_color="#FF5555", command=lambda: delete_command(self))
        self.del_btn.pack(side="right", padx=(5, 0))

    def toggle_visibility(self):
        self.is_hidden = not self.is_hidden
        self.key_entry.configure(show="•" if self.is_hidden else "")
        self.eye_btn.configure(fg_color="#555555" if self.is_hidden else "#2CC985")

    def get_data(self):
        return self.name_entry.get().strip(), self.key_entry.get().strip()

class AdvancedKeyWindow(ctk.CTkToplevel):
    def __init__(self, parent, current_text, on_save_callback):
        super().__init__(parent)
        self.title("⚙️ 進階 API Key 管理")
        self.geometry("620x500") 
        self.resizable(False, True)
        self.transient(parent)
        self.grab_set()
        self.focus_set()
        self.on_save_callback = on_save_callback
        self.rows = [] 
        ctk.CTkLabel(self, text="🔑 API Key 清單管理", font=("微軟正黑體", 18, "bold")).pack(pady=(15, 5))
        self.scroll_frame = ctk.CTkScrollableFrame(self, label_text="API 列表")
        self.scroll_frame.pack(fill="both", expand=True, padx=20, pady=5)
        self.parse_initial_data(current_text)
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=15)
        self.add_btn = ctk.CTkButton(btn_frame, text="➕ 新增一組 API", fg_color="#2CC985", command=lambda: self.add_row())
        self.add_btn.pack(side="left")
        ctk.CTkButton(btn_frame, text="確認儲存", width=100, command=self.save_and_close).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame, text="取消", width=80, fg_color="gray", command=self.destroy).pack(side="right", padx=5)

    def parse_initial_data(self, text):
        has_data = False
        if text:
            for line in text.split('\n'):
                line = line.strip()
                if not line: continue
                has_data = True
                if ":" in line:
                    name, key = line.split(":", 1)
                    self.add_row(name.strip(), key.strip())
                else:
                    self.add_row("", line.strip())
        if not has_data: self.add_row()

    def add_row(self, name="", key=""):
        row = APIKeyRow(self.scroll_frame, name, key, self.delete_row)
        self.rows.append(row)

    def delete_row(self, row_widget):
        if row_widget in self.rows:
            self.rows.remove(row_widget)
            row_widget.destroy()

    def save_and_close(self):
        result_lines = []
        auto_count = 1 
        for row in self.rows:
            name, key = row.get_data()
            if key: 
                if not name:
                    name = f"Key_{auto_count}"
                    auto_count += 1
                result_lines.append(f"{name}:{key}")
        self.on_save_callback("\n".join(result_lines))
        self.destroy()

class AIRefineWindow(ctk.CTkToplevel):
    def __init__(self, parent, current_rule, app_instance):
        super().__init__(parent)
        self.app_instance = app_instance
        self.title("✨ AI 規則優化 (Prompt Tuning)")
        self.geometry("800x650")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.focus_set()
        self.current_rule = current_rule

        # Layout Setup
        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Left side: User inputs (Bad & Expected)
        input_frame = ctk.CTkFrame(top_frame, fg_color="transparent")
        input_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        ctk.CTkLabel(input_frame, text="😔 AI 原本錯誤的輸出:", font=("微軟正黑體", 14, "bold")).pack(anchor="w")
        self.bad_text = ctk.CTkTextbox(input_frame, font=("微軟正黑體", 14), wrap="word")
        self.bad_text.pack(fill="both", expand=True, pady=(5, 10))
        
        ctk.CTkLabel(input_frame, text="✅ 您期望的理想輸出:", font=("微軟正黑體", 14, "bold")).pack(anchor="w")
        self.good_text = ctk.CTkTextbox(input_frame, font=("微軟正黑體", 14), wrap="word")
        self.good_text.pack(fill="both", expand=True, pady=(5, 10))

        # Right side: AI rule generation result
        result_frame = ctk.CTkFrame(top_frame, fg_color="transparent")
        result_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))

        ctk.CTkLabel(result_frame, text="💡 AI 計算出的新規則:", font=("微軟正黑體", 14, "bold")).pack(anchor="w")
        self.result_text = ctk.CTkTextbox(result_frame, font=("微軟正黑體", 14), wrap="word")
        self.result_text.pack(fill="both", expand=True, pady=(5, 10))
        
        # Controls
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=10)

        self.gen_btn = ctk.CTkButton(btn_frame, text="✨ 產生新規則", font=("微軟正黑體", 14, "bold"), fg_color="#8A2BE2", command=self.start_generation)
        self.gen_btn.pack(side="left", padx=5)

        self.apply_btn = ctk.CTkButton(btn_frame, text="套用至系統 (複製)", font=("微軟正黑體", 14), command=self.apply_rule)
        self.apply_btn.pack(side="left", padx=5)

        ctk.CTkButton(btn_frame, text="關閉", width=80, fg_color="gray", command=self.destroy).pack(side="right", padx=5)

    def start_generation(self):
        upper = self.bad_text.get("1.0", "end").strip()
        lower = self.good_text.get("1.0", "end").strip()
        if not upper or not lower:
            messagebox.showwarning("提示", "請確實填寫「原本輸出」與「期望輸出」！")
            return
            
        self.gen_btn.configure(state="disabled", text="處理中...")
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", "🧠 AI 正在思考並精煉提示詞，請稍候...")

        prompt = f"""你是一個「規則更新器」。
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
{upper}
------------------------------------
【使用者期望輸出 lower】
{lower}
------------------------------------
【舊規則 rule_text】
{self.current_rule}
------------------------------------
請輸出：【更新後的新規則】"""

        import threading
        threading.Thread(target=self._run_ai, args=(prompt,), daemon=True).start()

    def _run_ai(self, prompt):
        try:
            keys = self.app_instance.get_api_key()
            model = self.app_instance.get_model()
            
            if "gpt" in model.lower():
                from gpt import setup_gpt, changeRule
                setup_gpt(keys, model, "", log_callback=None)
            else:
                from gemini import setup_gemini, changeRule
                setup_gemini(keys, model, "", log_callback=None)

            new_rule = changeRule(prompt)
            if not new_rule: new_rule = "⚠️ API 未回應，請檢查金鑰或重新執行。"
            self.result_text.delete("1.0", "end")
            self.result_text.insert("1.0", new_rule)
        except Exception as e:
            self.result_text.delete("1.0", "end")
            self.result_text.insert("1.0", f"❌ 發生錯誤:\n{e}")
        finally:
            self.gen_btn.configure(state="normal", text="✨ 產生新規則")

    def apply_rule(self):
        new_r = self.result_text.get("1.0", "end").strip()
        if new_r and "發生錯誤" not in new_r and "AI 正在思考" not in new_r:
            self.clipboard_clear()
            self.clipboard_append(new_r)
            messagebox.showinfo("複製成功", "已經將 AI 新規則複製到剪貼簿，您可以手動貼至規則編輯視窗中！")

class RuleWindow(ctk.CTkToplevel):
    def __init__(self, parent, rule_text, on_save_callback, app_instance):
        super().__init__(parent)
        self.title("📝 編輯 OCR 規則")
        self.geometry("600x500")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.focus_set()
        self.on_save_callback = on_save_callback
        self.app_instance = app_instance

        ctk.CTkLabel(self, text="📝 OCR 辨識規則 (Prompt)", font=("微軟正黑體", 16, "bold")).pack(pady=(10, 5))
        self.rule_text_area = ctk.CTkTextbox(self, font=("微軟正黑體", 14), wrap="word")
        self.rule_text_area.pack(fill="both", expand=True, padx=20, pady=10)
        self.rule_text_area.insert("1.0", rule_text)

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=10)

        ctk.CTkButton(btn_frame, text="🔄 重置規則", fg_color="gray", hover_color="#555555", command=self.reset_rule).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="✨ AI 優化規則", fg_color="#8A2BE2", command=self.ai_refine_rule).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="儲存並關閉", width=120, command=self.save_and_close).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame, text="取消", width=80, fg_color="gray", command=self.destroy).pack(side="right", padx=5)

    def save_and_close(self):
        self.on_save_callback(self.rule_text_area.get("1.0", "end").strip())
        self.destroy()

    def reset_rule(self):
        self.rule_text_area.delete("1.0", "end")
        self.rule_text_area.insert("1.0", self.app_instance.default_rule_text)

    def ai_refine_rule(self):
        current_rule = self.rule_text_area.get("1.0", "end").strip()
        AIRefineWindow(self, current_rule, app_instance=self.app_instance)

class FileListItem(ctk.CTkFrame):
    def __init__(self, parent, file_path, delete_callback):
        super().__init__(parent, fg_color=("gray85", "gray25"), corner_radius=5)
        self.pack(fill="x", pady=5, padx=2)
        self.file_path = file_path
        filename = os.path.basename(file_path)
        self.is_pdf = filename.lower().endswith('.pdf')
        
        name_frame = ctk.CTkFrame(self, fg_color="transparent")
        name_frame.pack(fill="x", padx=5, pady=(5, 0))
        lbl = ctk.CTkLabel(name_frame, text=f"{'📄' if self.is_pdf else '🖼️'} {filename}", anchor="w", font=("微軟正黑體", 12, "bold"))
        lbl.pack(side="left", fill="x", expand=True)

        ctrl_frame = ctk.CTkFrame(self, fg_color="transparent")
        ctrl_frame.pack(fill="x", padx=5, pady=(2, 5))

        self.start_page = None
        self.end_page = None

        if self.is_pdf:
            page_frame = ctk.CTkFrame(ctrl_frame, fg_color="transparent")
            page_frame.pack(side="left")
            ctk.CTkLabel(page_frame, text="頁數:").pack(side="left", padx=2)
            self.start_page_entry = ctk.CTkEntry(page_frame, width=35, height=24, placeholder_text="1")
            self.start_page_entry.pack(side="left", padx=2)
            ctk.CTkLabel(page_frame, text="-").pack(side="left")
            total_pages_str = "末"
            try:
                import fitz
                pdf_doc = fitz.open(self.file_path)
                total_pages_str = str(len(pdf_doc))
                pdf_doc.close()
            except Exception:
                pass
                
            self.end_page_entry = ctk.CTkEntry(page_frame, width=40, height=24, placeholder_text=total_pages_str)
            self.end_page_entry.pack(side="left", padx=2)
            self.start_page = self.start_page_entry
            self.end_page = self.end_page_entry

        del_btn = ctk.CTkButton(ctrl_frame, text="❌ 移除", width=50, height=24, fg_color="#FF5555", command=lambda: delete_callback(self.file_path, self))
        del_btn.pack(side="right", padx=2)

    def get_individual_settings(self):
        if self.is_pdf and self.start_page and self.end_page:
            try: s = int(self.start_page.get()) if self.start_page.get() else None
            except: s = None
            try: e = int(self.end_page.get()) if self.end_page.get() else None
            except: e = None
            return s, e
        return None, None

class GeminiOCRApp:
    def __init__(self, root, default_rule_text=""):
        self.root = root
        self.default_rule_text = default_rule_text
        self.current_rule_text = default_rule_text
        self.root.title("Gemini OCR 自動化轉換工具 (CTK改版)")
        self.root.geometry("700x750")
        
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Create Tabview
        self.tabview = ctk.CTkTabview(root)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)
        self.tab_ocr = self.tabview.add("🔍 OCR 辨識")
        self.tab_word = self.tabview.add("📄 Word 排版優化")

        # ==========================================
        # 頁籤 1: OCR 辨識
        # ==========================================
        self.selected_file_paths = []
        self.file_items = []
        self.full_api_key_string = "" 

        # OCR 核心設定
        self.settings_frame = ctk.CTkFrame(self.tab_ocr, corner_radius=10)
        self.settings_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(self.settings_frame, text="🔧 核心設定", font=("微軟正黑體", 16, "bold")).pack(pady=(10, 5))

        top_row = ctk.CTkFrame(self.settings_frame, fg_color="transparent")
        top_row.pack(fill="x", padx=10, pady=2)
        
        ctk.CTkLabel(top_row, text="API Key 狀態: ", font=("微軟正黑體", 14, "bold")).pack(side="left", padx=5)
        self.api_status_label = ctk.CTkLabel(top_row, text="尚未設定", text_color="red", font=("微軟正黑體", 14))
        self.api_status_label.pack(side="left", padx=5)
        
        self.env_btn = ctk.CTkButton(top_row, text="🌐 設定環境變數", width=120, fg_color="#3B8ED0", command=self.open_env_window)
        self.env_btn.pack(side="right", padx=5)

        self.adv_btn = ctk.CTkButton(top_row, text="⚙️ 多組 Key", width=100, command=self.open_advanced_settings)
        self.adv_btn.pack(side="right", padx=5)
        
        self._update_api_status()

        bottom_row = ctk.CTkFrame(self.settings_frame, fg_color="transparent")
        bottom_row.pack(fill="x", padx=10, pady=2)

        ctk.CTkLabel(bottom_row, text="當前模型: ", font=("微軟正黑體", 14, "bold")).pack(side="left", padx=5)
        self.model_label = ctk.CTkLabel(bottom_row, text="gemini-2.5-flash", text_color="#2CC985", font=("微軟正黑體", 14))
        self.model_label.pack(side="left", padx=5)
        
        self.rule_btn = ctk.CTkButton(bottom_row, text="📝 編輯 OCR 規則", width=120, fg_color="#8A2BE2", command=self.open_rule_window)
        self.rule_btn.pack(side="right", padx=5)

        self.ignore_handwriting_var = ctk.BooleanVar(value=False)
        self.ignore_handwriting_cb = ctk.CTkCheckBox(bottom_row, text="忽略手寫字體", variable=self.ignore_handwriting_var, font=("微軟正黑體", 14))
        self.ignore_handwriting_cb.pack(side="right", padx=10)

        # OCR 工作區
        self.middle_container = ctk.CTkFrame(self.tab_ocr, fg_color="transparent")
        self.middle_container.pack(fill="both", expand=True, padx=10, pady=5)

        # OCR: 左側上傳區
        self.upload_frame = ctk.CTkFrame(self.middle_container, corner_radius=10, width=400)
        self.upload_frame.pack(side="left", fill="both", expand=False, padx=(0, 10))
        
        header_f = ctk.CTkFrame(self.upload_frame, fg_color="transparent")
        header_f.pack(fill="x", pady=(10, 5), padx=10)
        ctk.CTkLabel(header_f, text="📁 檔案列表", font=("微軟正黑體", 16, "bold")).pack(side="left")
        
        self.file_list_frame = ctk.CTkScrollableFrame(self.upload_frame, label_text="已選取 0 個檔案")
        self.file_list_frame.pack(fill="both", expand=True, padx=10, pady=5)

        page_range_frame = ctk.CTkFrame(self.upload_frame, fg_color="transparent")
        page_range_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(page_range_frame, text="全體 PDF 頁數:", font=("微軟正黑體", 12)).pack(side="left", padx=2)
        self.global_start_page = ctk.CTkEntry(page_range_frame, width=40, placeholder_text="1")
        self.global_start_page.pack(side="left", padx=2)
        ctk.CTkLabel(page_range_frame, text="-").pack(side="left")
        self.global_end_page = ctk.CTkEntry(page_range_frame, width=40, placeholder_text="全部")
        self.global_end_page.pack(side="left", padx=2)

        up_btn_frame = ctk.CTkFrame(self.upload_frame, fg_color="transparent")
        up_btn_frame.pack(fill="x", padx=10, pady=10)
        ctk.CTkButton(up_btn_frame, text="➕ 選擇檔案", fg_color="#3B8ED0", command=self.select_files).pack(side="left", fill="x", expand=True, padx=2)
        ctk.CTkButton(up_btn_frame, text="🗑️ 清空所有", width=80, fg_color="#FF5555", command=self.clear_files).pack(side="right", padx=2)

        # OCR: 右側輸出區
        self.output_frame = ctk.CTkFrame(self.middle_container, corner_radius=10)
        self.output_frame.pack(side="right", fill="both", expand=True)

        ctk.CTkLabel(self.output_frame, text="📄 即時編輯與輸出 (OUTPUT)", font=("微軟正黑體", 16, "bold")).pack(pady=(10, 5))
        self.output_text_area = ctk.CTkTextbox(self.output_frame, font=("微軟正黑體", 14), wrap="word")
        self.output_text_area.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # OCR: 執行區
        self.action_frame = ctk.CTkFrame(self.tab_ocr, fg_color="transparent")
        self.action_frame.pack(fill="x", padx=10, pady=5)

        self.start_btn = ctk.CTkButton(self.action_frame, text="開始執行轉換", font=("微軟正黑體", 16, "bold"), height=50, command=self.on_start_click)
        self.start_btn.pack(fill="x", pady=5)

        log_header_frame = ctk.CTkFrame(self.action_frame, fg_color="transparent")
        log_header_frame.pack(fill="x", pady=(5, 0))

        ctk.CTkLabel(log_header_frame, text="執行日誌:", font=("微軟正黑體", 12)).pack(side="left")
        ctk.CTkButton(log_header_frame, text="🗑️ 清空日誌", width=80, height=24, fg_color="#555555", command=self.clear_log).pack(side="right")

        self.log_area = ctk.CTkTextbox(self.action_frame, height=100, state="disabled", font=("微軟正黑體", 12))
        self.log_area.pack(fill="x", pady=(0, 5))

        # ==========================================
        # 頁籤 2: Word 排版優化 (toWord)
        # ==========================================
        self.word_selected_files = []
        
        title_frame = ctk.CTkFrame(self.tab_word, fg_color="transparent")
        title_frame.pack(fill="x", pady=(20, 10))
        ctk.CTkLabel(title_frame, text="📄 Word 文件排版優化器", font=("微軟正黑體", 20, "bold")).pack()
        ctk.CTkLabel(title_frame, text="將文件中文字型轉標楷體，英文數字符號轉 Times New Roman。\n並自動進行數學減號校正與特定英文字母斜體排版。", font=("微軟正黑體", 12), text_color="gray").pack(pady=(5, 0))
        
        self.word_list_frame = ctk.CTkScrollableFrame(self.tab_word, label_text="已選取 0 個 Word 檔案")
        self.word_list_frame.pack(fill="both", expand=True, padx=30, pady=10)
        
        word_btn_frame = ctk.CTkFrame(self.tab_word, fg_color="transparent")
        word_btn_frame.pack(fill="x", padx=30, pady=5)
        
        ctk.CTkButton(word_btn_frame, text="➕ 選擇檔案", width=120, command=self.select_word_files).pack(side="left")
        ctk.CTkButton(word_btn_frame, text="🗑️ 清空列表", width=100, fg_color="#FF5555", command=self.clear_word_files).pack(side="right")
        
        self.word_start_btn = ctk.CTkButton(self.tab_word, text="🚀 開始執行優化", font=("微軟正黑體", 16, "bold"), height=50, command=self.on_word_start)
        self.word_start_btn.pack(fill="x", padx=30, pady=(15, 5))
        
        self.word_log_area = ctk.CTkTextbox(self.tab_word, height=120, state="disabled", font=("微軟正黑體", 12))
        self.word_log_area.pack(fill="x", padx=30, pady=(5, 10))

    # ========================
    # 邏輯: Word 頁籤
    # ========================
    def word_log(self, message):
        self.word_log_area.configure(state="normal")
        self.word_log_area.insert("end", message + "\n")
        self.word_log_area.see("end")
        self.word_log_area.configure(state="disabled")

    def select_word_files(self):
        files = filedialog.askopenfilenames(
            title="請選擇 Word 檔案",
            filetypes=[("Word 檔案", "*.docx;*.doc")]
        )
        if files:
            for f in files:
                if f not in self.word_selected_files:
                    self.word_selected_files.append(f)
            self.update_word_list()

    def clear_word_files(self):
        self.word_selected_files = []
        self.update_word_list()

    def update_word_list(self):
        for widget in self.word_list_frame.winfo_children():
            widget.destroy()
        self.word_list_frame.configure(label_text=f"已選取 {len(self.word_selected_files)} 個 Word 檔案")
        for f in self.word_selected_files:
            ctk.CTkLabel(self.word_list_frame, text=f"📄 {os.path.basename(f)}", anchor="w").pack(fill="x", padx=5, pady=2)

    def on_word_start(self):
        if not self.word_selected_files:
            self.word_log("⚠️ 請先選擇 Word 檔案！")
            return
            
        self.word_start_btn.configure(state="disabled", text="⏳ 正在處理中...")
        threading.Thread(target=self.process_word_files, daemon=True).start()

    def process_word_files(self):
        self.word_log(f"--- 🚀 開始處理 {len(self.word_selected_files)} 個檔案 ---")
        for file in self.word_selected_files:
            filename = os.path.basename(file)
            self.word_log(f"🔄 正在處理: {filename} ...")
            try:
                changeWord(file)
                self.word_log(f"✅ {filename} 處理完成！")
            except Exception as e:
                self.word_log(f"❌ {filename} 發生錯誤: {e}")
        self.word_log("🎉 所有排版任務執行完畢！")
        self.word_start_btn.configure(state="normal", text="🚀 開始執行優化")

    # ========================
    # 邏輯: OCR 頁籤
    # ========================
    def _update_api_status(self):
        if self.full_api_key_string:
            valid_lines = [x for x in self.full_api_key_string.split('\n') if x.strip()]
            count = len(valid_lines)
            if count > 0:
                first_line = valid_lines[0]
                first_name = first_line.split(':')[0].strip() if ':' in first_line else "Key_1"
                if not first_name: first_name = "Key_1"
                self.api_status_label.configure(text=f"已連線 (優先使用：{first_name}，共 {count} 組)", text_color="#2CC985")
            return

        cached_env = os.environ.get("GEMINI_API_KEY")
        registry_exists = False
        try:
            reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_READ)
            try:
                value, _ = winreg.QueryValueEx(reg_key, "GEMINI_API_KEY")
                if value: registry_exists = True
                cached_env = value 
            except FileNotFoundError: pass
            winreg.CloseKey(reg_key)
        except Exception: pass

        if registry_exists or cached_env:
            key_show = cached_env[:8] + "..." + cached_env[-4:] if len(cached_env) > 12 else "已設定"
            self.api_status_label.configure(text=f"已連線 ({key_show})", text_color="#2CC985")
        else:
            self.api_status_label.configure(text="未設定 / 查無金鑰", text_color="red")

    def open_env_window(self):
        EnvVarWindow(self.root, self.on_env_saved)

    def on_env_saved(self):
        self._update_api_status()
        if os.environ.get("GEMINI_API_KEY") or self.full_api_key_string:
            self.log("✅ API Key 設定已更新。")
        else:
            self.log("ℹ️ API Key 環境變數已清除。")

    def open_advanced_settings(self):
        AdvancedKeyWindow(self.root, self.full_api_key_string, self.on_advanced_save)

    def on_advanced_save(self, text):
        self.full_api_key_string = text
        self._update_api_status()
        self.log("✅ 多組 API Key 設置已儲存。")


    def open_rule_window(self):
        RuleWindow(self.root, self.current_rule_text, self.on_rule_saved, app_instance=self)
        
    def on_rule_saved(self, text):
        self.current_rule_text = text
        self.log("✅ OCR 規則已保存。")

    def clear_log(self):
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="請選擇要處理的圖片或 PDF",
            filetypes=[("圖片與 PDF", "*.png;*.jpg;*.jpeg;*.pdf"), ("所有檔案", "*.*")]
        )
        if files:
            for f in files:
                if f not in self.selected_file_paths:
                    self.selected_file_paths.append(f)
                    item = FileListItem(self.file_list_frame, f, self.remove_single_file)
                    self.file_items.append(item)
            self.file_list_frame.configure(label_text=f"已選取 {len(self.selected_file_paths)} 個檔案")

    def remove_single_file(self, path, item_widget):
        if path in self.selected_file_paths:
            self.selected_file_paths.remove(path)
        if item_widget in self.file_items:
            self.file_items.remove(item_widget)
        item_widget.destroy()
        self.file_list_frame.configure(label_text=f"已選取 {len(self.selected_file_paths)} 個檔案")

    def clear_files(self):
        self.selected_file_paths = []
        for item in self.file_items:
            item.destroy()
        self.file_items = []
        self.file_list_frame.configure(label_text="已選取 0 個檔案")

    def on_start_click(self):
        self.log("🚀 開始處理...")
        self.log(f"📝 規則預覽 (前10字): {self.current_rule_text[:10]}...")
        gst = self.global_start_page.get()
        ged = self.global_end_page.get()
        self.log(f"🌐 全體 PDF 設定: {gst} 到 {ged}")
        
        for item in self.file_items:
            if item.is_pdf:
                s, e = item.get_individual_settings()
                self.log(f"📄 PDF [{os.path.basename(item.file_path)}] - 個別設定頁數: {s} 到 {e}")

    def log(self, message):
        self.log_area.configure(state="normal")
        self.log_area.insert("end", message + "\n")
        self.log_area.see("end")
        self.log_area.configure(state="disabled")

    def append_output(self, text):
        self.output_text_area.insert('end', text + '\n')
        self.output_text_area.see('end')

    def get_api_key(self): return self.full_api_key_string or os.environ.get("GEMINI_API_KEY", "")
    def get_model(self): return self.model_label.cget("text")
    def get_rule_text(self): 
        base_rule = self.current_rule_text
        if hasattr(self, 'ignore_handwriting_var') and self.ignore_handwriting_var.get():
            base_rule += "\n\n<additional_rule>\n請完全忽略圖片中的任何「手寫字體」、「手寫筆記」或「手寫塗鴉」，僅辨識並輸出原始的「印刷字體」內容。\n</additional_rule>"
        return base_rule
    def get_selected_files(self): return self.selected_file_paths
    
    def get_file_settings(self):
        try: gs = int(self.global_start_page.get()) if self.global_start_page.get() else 1
        except: gs = 1
        try: ge = int(self.global_end_page.get()) if self.global_end_page.get() else 0
        except: ge = 0
        
        settings = {}
        for item in self.file_items:
            path = item.file_path
            s, e = gs, ge
            if item.is_pdf:
                indy_s, indy_e = item.get_individual_settings()
                if indy_s is not None: s = indy_s
                if indy_e is not None: e = indy_e
            settings[path] = {"start": s, "end": e}
        return settings

if __name__ == "__main__":
    dummy_rule = "這是測試用的預設規則。\n1. 請保持原意\n2. 不要亂改"
    root = ctk.CTk()
    app = GeminiOCRApp(root, default_rule_text=dummy_rule)
    root.mainloop()
