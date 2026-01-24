import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
import winreg  # <--- 新增這行，用來查證 Windows 登錄檔正本

# ==========================================
# 自定義元件區
# ==========================================

class ModelInfoPopup(ctk.CTkToplevel):
    """ 自定義彈出視窗：模態視窗 (Modal) """
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

        self.ok_btn = ctk.CTkButton(self, text="我知道了", width=120, height=40, 
                                    font=("微軟正黑體", 14, "bold"),
                                    command=self.destroy)
        self.ok_btn.pack(side="bottom", pady=20)


class EnvVarWindow(ctk.CTkToplevel):
    """ 環境變數設定視窗 """
    def __init__(self, parent, on_save_callback):
        super().__init__(parent)
        self.title("🌐 設定環境變數 API Key")
        self.geometry("500x250")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.focus_set()
        self.on_save_callback = on_save_callback

        # 標題
        ctk.CTkLabel(self, text="設定系統環境變數 (GEMINI_API_KEY)", font=("微軟正黑體", 16, "bold")).pack(pady=(20, 10))
        ctk.CTkLabel(self, text="留空並儲存即可「完全刪除」此變數。", font=("微軟正黑體", 12), text_color="gray").pack(pady=(0, 20))

        # 輸入區塊
        input_frame = ctk.CTkFrame(self, fg_color="transparent")
        input_frame.pack(fill="x", padx=20)

        self.is_hidden = True
        
        # 嘗試讀取現有的環境變數
        current_env = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or ""

        self.key_entry = ctk.CTkEntry(input_frame, width=350, placeholder_text="請貼上 API Key (留空則刪除)", show="•")
        self.key_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.key_entry.insert(0, current_env)

        self.eye_btn = ctk.CTkButton(input_frame, text="👁", width=40, fg_color="#555555", hover_color="#777777",
                                     command=self.toggle_visibility)
        self.eye_btn.pack(side="left", padx=5)

        # 按鈕區
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=30)

        ctk.CTkButton(btn_frame, text="確認設定", width=120, command=self.save_env).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame, text="取消", width=80, fg_color="gray", command=self.destroy).pack(side="right", padx=5)

    def toggle_visibility(self):
        self.is_hidden = not self.is_hidden
        if self.is_hidden:
            self.key_entry.configure(show="•")
            self.eye_btn.configure(fg_color="#555555")
        else:
            self.key_entry.configure(show="")
            self.eye_btn.configure(fg_color="#2CC985")

    def save_env(self):
        """ 儲存或刪除環境變數的核心邏輯 """
        key = self.key_entry.get().strip()
        
        # --- 情況 A: 使用者留空 -> 執行完全刪除 ---
        if not key:
            # 1. 移除當前 Python 執行環境的變數
            if "GEMINI_API_KEY" in os.environ:
                del os.environ["GEMINI_API_KEY"]
            
            # 2. (可選) 如果之前有產生 .env，順便刪掉它，保持乾淨
            if os.path.exists(".env"):
                try:
                    os.remove(".env")
                except Exception:
                    pass

            # 3. 移除 Windows 系統使用者環境變數
            try:
                subprocess.run(
                    r'REG DELETE "HKCU\Environment" /V GEMINI_API_KEY /F', 
                    shell=True, 
                    stdout=subprocess.DEVNULL, 
                    stderr=subprocess.DEVNULL
                )
            except Exception:
                pass
            
            messagebox.showinfo("提示", "已清除環境變數。")

        # --- 情況 B: 使用者有輸入 -> 只寫入系統變數 ---
        else:
            # 1. 設定當前環境 (立即生效)
            os.environ["GEMINI_API_KEY"] = key
            
            # (原來的 "寫入 .env" 區塊已被刪除)

            # 2. 寫入 Windows 系統使用者環境變數 (永久生效)
            try:
                subprocess.run(
                    f'setx GEMINI_API_KEY "{key}"', 
                    shell=True, 
                    stdout=subprocess.DEVNULL, 
                    stderr=subprocess.DEVNULL
                )
            except Exception:
                pass

        # 回調通知主視窗更新介面
        self.on_save_callback()
        self.destroy()

class APIKeyRow(ctk.CTkFrame):
    """ 單一列元件 """
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

        self.eye_btn = ctk.CTkButton(self, text="👁", width=30, fg_color="#555555", hover_color="#777777",
                                     command=self.toggle_visibility)
        self.eye_btn.pack(side="left", padx=2)
            
        self.del_btn = ctk.CTkButton(self, text="❌", width=30, fg_color="#FF5555", hover_color="#CC0000",
                                     command=lambda: delete_command(self))
        self.del_btn.pack(side="right", padx=(5, 0))

    def toggle_visibility(self):
        self.is_hidden = not self.is_hidden
        if self.is_hidden:
            self.key_entry.configure(show="•")
            self.eye_btn.configure(fg_color="#555555")
        else:
            self.key_entry.configure(show="")
            self.eye_btn.configure(fg_color="#2CC985")

    def get_data(self):
        return self.name_entry.get().strip(), self.key_entry.get().strip()


class AdvancedKeyWindow(ctk.CTkToplevel):
    """ 彈出視窗：圖形化管理多組 API Key """
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
        ctk.CTkLabel(self, text="若名稱留空，系統將自動命名為 Key_1, Key_2...", font=("微軟正黑體", 12), text_color="gray").pack(pady=(0, 10))

        self.scroll_frame = ctk.CTkScrollableFrame(self, label_text="API 列表")
        self.scroll_frame.pack(fill="both", expand=True, padx=20, pady=5)

        self.parse_initial_data(current_text)

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=15)

        self.add_btn = ctk.CTkButton(btn_frame, text="➕ 新增一組 API", fg_color="#2CC985", hover_color="#00A86B",
                                     command=lambda: self.add_row())
        self.add_btn.pack(side="left")

        ctk.CTkButton(btn_frame, text="確認儲存", width=100, command=self.save_and_close).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame, text="取消", width=80, fg_color="gray", command=self.destroy).pack(side="right", padx=5)

    def parse_initial_data(self, text):
        lines = text.split('\n')
        has_data = False
        for line in lines:
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
        
        final_string = "\n".join(result_lines)
        self.on_save_callback(final_string)
        self.destroy()

# ==========================================
# 主程式區
# ==========================================

class GeminiOCRApp:
    def __init__(self, root, default_rule_text=""):
        self.root = root
        self.default_rule_text = default_rule_text 
        self.root.title("Gemini OCR 自動化轉換工具 (CTK版)")
        self.root.geometry("750x800")
        
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.full_api_key_string = "" 
        self.selected_file_paths = [] 

        # --- 區塊 1：核心設定 ---
        self.settings_frame = ctk.CTkFrame(root, corner_radius=10)
        self.settings_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(self.settings_frame, text="🔧 核心設定", font=("微軟正黑體", 16, "bold")).pack(pady=(10, 5))

        # API Key
        key_frame = ctk.CTkFrame(self.settings_frame, fg_color="transparent")
        key_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(key_frame, text="API Key:", font=("微軟正黑體", 14)).pack(side="left", padx=5)
        
        self.api_key_entry = ctk.CTkEntry(key_frame, width=200)
        self.api_key_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        # 【關鍵】初始化檢查
        self._update_placeholder_status()
        self.api_key_entry.configure(state="readonly")
        
        # 按鈕容器
        btn_container = ctk.CTkFrame(key_frame, fg_color="transparent")
        btn_container.pack(side="left", padx=5)

        self.adv_btn = ctk.CTkButton(btn_container, text="⚙️ 多組 Key", width=100, 
                                     command=self.open_advanced_settings)
        self.adv_btn.pack(side="left", padx=2)

        self.env_btn = ctk.CTkButton(btn_container, text="🌐 設定環境變數", width=100, 
                                     fg_color="#3B8ED0", 
                                     command=self.open_env_window)
        self.env_btn.pack(side="left", padx=2)

        # 模型選擇
        model_frame = ctk.CTkFrame(self.settings_frame, fg_color="transparent")
        model_frame.pack(fill="x", padx=10, pady=(5, 15))
        ctk.CTkLabel(model_frame, text="Gemini 模型:", font=("微軟正黑體", 14)).pack(side="left", padx=5)
        self.model_combobox = ctk.CTkComboBox(model_frame, values=["gemini-2.5-flash", "gemini-1.5-flash", "gemini-1.5-pro"], width=200, state="readonly", command=self.on_model_change)
        self.model_combobox.set("gemini-2.5-flash")
        self.model_combobox.pack(side="left", padx=5)

        # ==========================================
        # 區塊 2：工作區 (左：檔案上傳 | 右：規則編輯)
        # ==========================================
        self.middle_container = ctk.CTkFrame(root, fg_color="transparent")
        self.middle_container.pack(fill="both", expand=True, padx=20, pady=5)

        # --- 左側：檔案上傳區 ---
        self.upload_frame = ctk.CTkFrame(self.middle_container, corner_radius=10, width=250)
        self.upload_frame.pack(side="left", fill="both", padx=(0, 10))
        
        ctk.CTkLabel(self.upload_frame, text="📁 檔案列表", font=("微軟正黑體", 16, "bold")).pack(pady=(10, 5))
        self.file_list_frame = ctk.CTkScrollableFrame(self.upload_frame, label_text="已選取 0 個檔案")
        self.file_list_frame.pack(fill="both", expand=True, padx=10, pady=5)

        up_btn_frame = ctk.CTkFrame(self.upload_frame, fg_color="transparent")
        up_btn_frame.pack(fill="x", padx=10, pady=10)
        ctk.CTkButton(up_btn_frame, text="➕ 選擇檔案", fg_color="#3B8ED0", command=self.select_files).pack(side="left", fill="x", expand=True, padx=2)
        ctk.CTkButton(up_btn_frame, text="🗑️", width=40, fg_color="#FF5555", hover_color="#CC0000", command=self.clear_files).pack(side="right", padx=2)


        # --- 右側：規則編輯區 ---
        self.rule_frame = ctk.CTkFrame(self.middle_container, corner_radius=10)
        self.rule_frame.pack(side="right", fill="both", expand=True)

        ctk.CTkLabel(self.rule_frame, text="📝 OCR 辨識規則 (Prompt)", font=("微軟正黑體", 16, "bold")).pack(pady=(10, 5))
        self.rule_text_area = ctk.CTkTextbox(self.rule_frame, font=("微軟正黑體", 14), wrap="word")
        self.rule_text_area.pack(fill="both", expand=True, padx=10, pady=(0, 5))
        
        self.rule_text_area.insert("1.0", self.default_rule_text)

        rule_btn_frame = ctk.CTkFrame(self.rule_frame, fg_color="transparent")
        rule_btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(rule_btn_frame, text="🔄 重置規則", fg_color="gray", hover_color="#555555",
                      command=self.reset_rule).pack(side="left", expand=True, fill="x", padx=(0, 5))
        ctk.CTkButton(rule_btn_frame, text="✨ AI 優化規則", fg_color="#8A2BE2", hover_color="#7B1FA2",
                      command=self.ai_refine_rule).pack(side="right", expand=True, fill="x", padx=(5, 0))


        # ==========================================
        # 區塊 3：執行與日誌區
        # ==========================================
        self.action_frame = ctk.CTkFrame(root, fg_color="transparent")
        self.action_frame.pack(fill="x", padx=20, pady=10)

        self.start_btn = ctk.CTkButton(self.action_frame, 
                                       text="儲存設定並開始執行", 
                                       font=("微軟正黑體", 16, "bold"),
                                       height=50,
                                       command=self.on_start_click)
        self.start_btn.pack(fill="x", pady=5)

        log_header_frame = ctk.CTkFrame(self.action_frame, fg_color="transparent")
        log_header_frame.pack(fill="x", pady=(10, 0))

        ctk.CTkLabel(log_header_frame, text="執行日誌:", font=("微軟正黑體", 12)).pack(side="left")
        
        ctk.CTkButton(log_header_frame, text="🗑️ 清空日誌", width=80, height=24,
                      fg_color="#555555", hover_color="#333333",
                      font=("微軟正黑體", 12),
                      command=self.clear_log).pack(side="right")

        self.log_area = ctk.CTkTextbox(self.action_frame, height=120, state="disabled", font=("微軟正黑體", 12))
        self.log_area.pack(fill="x", pady=(0, 10))

    # ----------------------------------------------------------------
    # 邏輯方法
    # ----------------------------------------------------------------
    
    # 【Helper】檢查環境變數並更新 Placeholder
# 【核心 Helper】檢查環境變數 (雙重驗證版)
    def _update_placeholder_status(self):
        # 1. 先看快取 (os.environ)
        cached_env = os.environ.get("GEMINI_API_KEY")
        
        # 2. 再看 Windows 登錄檔正本 (Registry)
        registry_exists = False
        try:
            # 打開 HKEY_CURRENT_USER\Environment
            reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_READ)
            try:
                # 嘗試讀取 GEMINI_API_KEY
                value, reg_type = winreg.QueryValueEx(reg_key, "GEMINI_API_KEY")
                if value:
                    registry_exists = True
            except FileNotFoundError:
                pass # 找不到就是真的沒有
            winreg.CloseKey(reg_key)
        except Exception:
            pass

        # 3. 綜合判斷：只有當「正本」存在，才算真的有
        # 這樣可以避免 VS Code 的幽靈快取騙我們
        if registry_exists:
            self.api_key_entry.configure(placeholder_text="已偵測到環境變數 (系統)")
        elif cached_env:
            # 這種情況就是「幽靈」：系統刪了，但終端機還留著
            # 我們選擇顯示「尚未設定」，並順手把當前程式的假變數清掉
            del os.environ["GEMINI_API_KEY"]
            self.api_key_entry.configure(placeholder_text="尚未設定")
        else:
            self.api_key_entry.configure(placeholder_text="尚未設定")

    def open_env_window(self):
        EnvVarWindow(self.root, self.on_env_saved)

    def on_env_saved(self):
        # 檢查環境變數狀態
        has_env = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
        
        if has_env:
            self.log("✅ 環境變數 GEMINI_API_KEY 已更新。")
        else:
            self.log("ℹ️ 環境變數 GEMINI_API_KEY 已移除。")

        # 更新文字框狀態
        if not self.full_api_key_string:
            self.api_key_entry.configure(state="normal")
            self.api_key_entry.delete(0, "end")
            self._update_placeholder_status() # 呼叫 Helper
            self.api_key_entry.configure(state="readonly")

    def clear_log(self):
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")

    def reset_rule(self):
        self.rule_text_area.delete("1.0", "end")
        self.rule_text_area.insert("1.0", self.default_rule_text)
        self.log("ℹ️ 規則已重置為系統預設值。")

    def ai_refine_rule(self):
        msg = "🚧 功能開發中 🚧\n\n敬請期待後續更新！"
        ModelInfoPopup(self.root, "功能尚未開放", msg)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="請選擇要處理的圖片或 PDF",
            filetypes=[("支援的檔案", "*.png;*.jpg;*.jpeg;*.pdf;*.docx;*.doc"), ("所有檔案", "*.*")]
        )
        if files:
            for f in files:
                if f not in self.selected_file_paths:
                    self.selected_file_paths.append(f)
            self.update_file_list_display()

    def clear_files(self):
        self.selected_file_paths = []
        self.update_file_list_display()

    def update_file_list_display(self):
        for widget in self.file_list_frame.winfo_children():
            widget.destroy()
        count = len(self.selected_file_paths)
        self.file_list_frame.configure(label_text=f"已選取 {count} 個檔案")
        for path in self.selected_file_paths:
            filename = os.path.basename(path)
            lbl = ctk.CTkLabel(self.file_list_frame, text=f"📄 {filename}", anchor="w")
            lbl.pack(fill="x", pady=2)

    def on_model_change(self, choice):
        descriptions = {
            "gemini-2.5-flash": "🚀 最新一代平衡型 (推薦)。",
            "gemini-1.5-flash": "⚡ 輕量經濟型。",
            "gemini-1.5-pro": "🧠 高階推理型。"
        }
        info_text = descriptions.get(choice, "目前沒有詳細資訊。")
        ModelInfoPopup(self.root, title=f"關於 {choice}", message=info_text)

    def open_advanced_settings(self):
        current_val = self.full_api_key_string 
        AdvancedKeyWindow(self.root, current_val, self.on_advanced_save)

    def on_advanced_save(self, text):
        self.full_api_key_string = text
        count = len([x for x in text.split('\n') if x.strip()])
        self.api_key_entry.configure(state="normal")
        self.api_key_entry.delete(0, "end")
        if count == 0:
            self.api_key_entry.insert(0, "")
            # 如果進階 Key 為空，回退顯示環境變數狀態
            self._update_placeholder_status()
        else:
            first_key = text.split('\n')[0].split(':')[-1].strip()
            safe_display = f"{first_key[:10]}... (共 {count} 組)"
            self.api_key_entry.insert(0, safe_display)
            self.log(f"✅ 設定更新：已載入 {count} 組 Key。")
        self.api_key_entry.configure(state="readonly")

    def on_start_click(self):
        print(f"檔案: {self.selected_file_paths}")
        print(f"Key: {self.full_api_key_string}")
        print("準備開始...")

    def log(self, message):
        self.log_area.configure(state="normal")
        self.log_area.insert("end", message + "\n")
        self.log_area.see("end")
        self.log_area.configure(state="disabled")

    def get_api_key(self): return self.full_api_key_string
    def get_model(self): return self.model_combobox.get()
    def get_rule_text(self): return self.rule_text_area.get("1.0", "end").strip()
    def get_selected_files(self): return self.selected_file_paths


if __name__ == "__main__":
    dummy_rule = "這是測試用的預設規則。\n1. 請保持原意\n2. 不要亂改"
    root = ctk.CTk()
    app = GeminiOCRApp(root, default_rule_text=dummy_rule)
    root.mainloop()