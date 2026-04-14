import os
import sys
import traceback
from typing import Union, List, Dict, Optional, Callable

from google import genai
from google.genai import types
from google.genai import errors
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception

# ==========================================
# 1. 常數設定
# ==========================================
ROTATE_TRIGGER_KEYWORDS = [
    "429",
    "RESOURCE_EXHAUSTED",
    "API key not valid",
    "PERMISSION_DENIED"
]

ERROR_HINT_MAP = {
    "429": "今日 Gemini API 免費額度已用盡。",
    "RESOURCE_EXHAUSTED": "今日 Gemini API 免費額度已用盡。",
    "API key not valid": "API Key 無效，請檢查設定。",
    "PERMISSION_DENIED": "API Key 權限不足或已失效。",
    "404": "找不到指定的模型（請檢查模型名稱是否正確）。",
    "INVALID_ARGUMENT": "傳送的參數格式有誤（例如圖片太大或文字不合法）。"
}

# ==========================================
# 2. Gemini 處理器類別 (執行緒安全)
# ==========================================
class GeminiProcessor:
    def __init__(self, key_content: Union[str, List[str], Dict[str, str]], 
                 model_name: str, 
                 rule_text: str,
                 log_callback: Optional[Callable[[str], None]] = None):
        
        self.gui_log = log_callback 
        self.api_key_dict = {}
        self.model = model_name
        self.rule = rule_text
        self.client = None
        self.current_key_name = None

        # 解析 Key
        if key_content:
            if isinstance(key_content, dict):
                self.api_key_dict = key_content.copy()
            elif isinstance(key_content, list):
                self.api_key_dict = {f"Key_{i+1}": k for i, k in enumerate(key_content)}
            elif isinstance(key_content, str):
                lines = key_content.strip().split('\n')
                for line in lines:
                    line = line.strip()
                    if not line: continue
                    if ":" in line:
                        name, key = line.split(":", 1)
                        self.api_key_dict[name.strip()] = key.strip()
                    else:
                        self.api_key_dict[f"Key_{len(self.api_key_dict)+1}"] = line.strip()

        env_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
        if env_key:
            self.api_key_dict["環境變數"] = env_key
            
        if not self.api_key_dict:
            self.log_msg("❌ 錯誤：未提供 API Key，且系統環境變數未設定。")
            raise ValueError("No API Key provided")

        self.current_key_name = list(self.api_key_dict.keys())[0]
        self._init_client_with_current_key()

    def log_msg(self, msg: str):
        print(msg) 
        if self.gui_log:
            self.gui_log(msg)

    def _init_client_with_current_key(self):
        if not self.current_key_name: return
        try:
            key_value = self.api_key_dict[self.current_key_name].strip()
            self.client = genai.Client(api_key=key_value)
            self.log_msg(f"🔑 初始化成功，目前使用【{self.current_key_name}】")
        except Exception as e:
            self.log_msg(f"❌ Client 初始化失敗 ({self.current_key_name})")

    def rotate_key(self) -> bool:
        all_keys = list(self.api_key_dict.keys())
        try:
            current_index = all_keys.index(self.current_key_name)
        except ValueError:
            current_index = -1

        if current_index + 1 < len(all_keys):
            self.current_key_name = all_keys[current_index + 1]
            self._init_client_with_current_key()
            return True
        else:
            self.log_msg(f"😭 所有 API Key ({len(all_keys)} 組) 皆已耗盡！")
            return False

    def handle_fatal_error(self, err_msg: str):
        for key, hint in ERROR_HINT_MAP.items():
            if key in err_msg:
                self.log_msg(f"\n{'!'*30}\n【執行停止】{hint}\n{'!'*30}\n")
                raise RuntimeError(f"Fatal Error: {hint}")
        self.log_msg(f"🚨 發生未預期的錯誤: {err_msg}")
        raise RuntimeError(f"Fatal Error: {err_msg}")

    @retry(
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=2, min=10, max=120),
        retry=retry_if_exception(lambda e: isinstance(e, errors.ServerError)),
        reraise=True
    )
    def changeRule(self, prompt: str) -> str:
        if self.client is None:
            self.log_msg("❌ 錯誤：Gemini 尚未初始化")
            return ""
        while True:
            try:
                return self.client.models.generate_content(
                    model=self.model,
                    contents=[prompt]
                ).text
            except errors.ClientError as e:
                err_msg = str(e)
                should_rotate = any(keyword in err_msg for keyword in ROTATE_TRIGGER_KEYWORDS)
                if should_rotate:
                    if self.rotate_key(): continue 
                    else: raise RuntimeError("All API Keys exhausted")
                self.handle_fatal_error(err_msg)
            except Exception as e:
                self.log_msg(f"🚨 未預期錯誤: {e}")
                raise e

    @retry(
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=2, min=10, max=120),
        retry=retry_if_exception(lambda e: isinstance(e, errors.ServerError)),
        reraise=True
    )
    def gemini_identify(self, pic_path: str) -> str:
        if self.client is None:
            self.log_msg("❌ 錯誤：Gemini 尚未初始化")
            return ""

        try:
            with open(pic_path, 'rb') as f:
                image_bytes = f.read()
        except Exception as e:
            self.log_msg(f"❌ 無法讀取圖片: {pic_path}")
            return ""

        while True:
            try:
                return self.client.models.generate_content(
                    model=self.model,
                    contents=[
                        types.Part.from_bytes(data=image_bytes, mime_type='image/png'),
                        self.rule
                    ]
                ).text
            except errors.ClientError as e:
                err_msg = str(e)
                should_rotate = any(keyword in err_msg for keyword in ROTATE_TRIGGER_KEYWORDS)
                if should_rotate:
                    hint_message = "API 連線異常"
                    for k, h in ERROR_HINT_MAP.items():
                        if k in err_msg:
                            hint_message = h
                            break
                    self.log_msg(f"⚠️ {hint_message}")
                    self.log_msg("🔄 嘗試切換 Key...")
                    if self.rotate_key():
                        continue 
                    else:
                        raise RuntimeError("All API Keys exhausted")
                self.handle_fatal_error(err_msg)
            except Exception as e:
                self.log_msg(f"🚨 未預期錯誤: {e}")
                raise e

# ==========================================
# 3. 為了相容性保留的全局介面 (但不建議在多執行緒併發使用)
# ==========================================
_default_processor = None

def setup_gemini(key_content, model_name, rule_text, log_callback=None):
    global _default_processor
    _default_processor = GeminiProcessor(key_content, model_name, rule_text, log_callback)

def gemini_identify(pic_path):
    if _default_processor:
        return _default_processor.gemini_identify(pic_path)
    return ""

def changeRule(prompt):
    if _default_processor:
        return _default_processor.changeRule(prompt)
    return ""