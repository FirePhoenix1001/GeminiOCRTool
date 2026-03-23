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
# 2. 全域變數
# ==========================================
client: Optional[genai.Client] = None
model: Optional[str] = None
rule: Optional[str] = None

api_key_dict: Dict[str, str] = {}
current_key_name: Optional[str] = None
gui_log: Optional[Callable[[str], None]] = None

def log_msg(msg: str):
    """ 統一輸出介面：同時印在終端機與 GUI """
    print(msg) 
    if gui_log:
        gui_log(msg)

# ==========================================
# 3. 初始化邏輯 (修正版)
# ==========================================
def setup_gemini(key_content: Union[str, List[str], Dict[str, str]], 
                 model_name: str, 
                 rule_text: str,
                 log_callback: Optional[Callable[[str], None]] = None):
    
    global client, model, rule, api_key_dict, current_key_name, gui_log
    
    gui_log = log_callback 
    api_key_dict = {}

    # --- 【關鍵修正開始】 讓 gemini.py 看懂 GUI 傳來的字串 ---
    if key_content:
        # 情況 A: 傳入的是字典 (CLI 模式)
        if isinstance(key_content, dict):
            api_key_dict = key_content.copy()
            
        # 情況 B: 傳入的是列表
        elif isinstance(key_content, list):
            api_key_dict = {f"Key_{i+1}": k for i, k in enumerate(key_content)}
            
        # 情況 C: 傳入的是字串 (GUI 模式) -> 需要解析 "Name:Key" 格式
        elif isinstance(key_content, str):
            lines = key_content.strip().split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue
                
                if ":" in line:
                    # 如果格式是 "名稱:Key"
                    name, key = line.split(":", 1)
                    api_key_dict[name.strip()] = key.strip()
                else:
                    # 如果只有 Key
                    api_key_dict[f"Key_{len(api_key_dict)+1}"] = line.strip()
    # --- 【關鍵修正結束】 ---

    # 2. 處理環境變數
    env_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if env_key:
        backup_name = "環境變數"
        api_key_dict[backup_name] = env_key
        
    # 3. 防呆
    if not api_key_dict:
        log_msg("❌ 錯誤：未提供 API Key，且系統環境變數未設定。")
        raise ValueError("No API Key provided")

    model = model_name
    rule = rule_text
    
    current_key_name = list(api_key_dict.keys())[0]
    _init_client_with_current_key()

def _init_client_with_current_key():
    global client
    if not current_key_name: return

    try:
        key_value = api_key_dict[current_key_name].strip()
        client = genai.Client(api_key=key_value)
        log_msg(f"🔑 初始化成功，目前使用【{current_key_name}】")
    except Exception as e:
        log_msg(f"❌ Client 初始化失敗 ({current_key_name})")

def rotate_key() -> bool:
    global current_key_name
    all_keys = list(api_key_dict.keys())
    try:
        current_index = all_keys.index(current_key_name)
    except ValueError:
        current_index = -1

    if current_index + 1 < len(all_keys):
        next_key_name = all_keys[current_index + 1]
        current_key_name = next_key_name
        _init_client_with_current_key()
        return True
    else:
        log_msg(f"😭 所有 API Key ({len(all_keys)} 組) 皆已耗盡！")
        return False

# ==========================================
# 4. 錯誤處理
# ==========================================
def handle_fatal_error(err_msg: str):
    for key, hint in ERROR_HINT_MAP.items():
        if key in err_msg:
            log_msg(f"\n{'!'*30}\n【執行停止】{hint}\n{'!'*30}\n")
            raise RuntimeError(f"Fatal Error: {hint}")
    
    log_msg(f"🚨 發生未預期的錯誤: {err_msg}")
    raise RuntimeError(f"Fatal Error: {err_msg}")

def should_retry(exception):
    return isinstance(exception, errors.ServerError)

# ==========================================
# 5. 主功能函式
# ==========================================
@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=2, min=10, max=120),
    retry=retry_if_exception(should_retry),
    reraise=True
)
def changeRule(prompt: str) -> str:
    if client is None:
        log_msg("❌ 錯誤：Gemini 尚未初始化")
        return ""
    while True:
        try:
            return client.models.generate_content(
                model=model,
                contents=[prompt]
            ).text
        except errors.ClientError as e:
            err_msg = str(e)
            should_rotate = any(keyword in err_msg for keyword in ROTATE_TRIGGER_KEYWORDS)
            if should_rotate:
                if rotate_key(): continue 
                else: raise RuntimeError("All API Keys exhausted")
            handle_fatal_error(err_msg)
        except Exception as e:
            log_msg(f"🚨 未預期錯誤: {e}")
            raise e

@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=2, min=10, max=120),
    retry=retry_if_exception(should_retry),
    reraise=True
)
def gemini_identify(pic_path: str) -> str:
    if client is None:
        log_msg("❌ 錯誤：Gemini 尚未初始化")
        return ""

    try:
        with open(pic_path, 'rb') as f:
            image_bytes = f.read()
    except Exception as e:
        log_msg(f"❌ 無法讀取圖片: {pic_path}")
        return ""

    while True:
        try:
            return client.models.generate_content(
                model=model,
                contents=[
                    types.Part.from_bytes(data=image_bytes, mime_type='image/png'),
                    rule
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
                
                log_msg(f"⚠️ {hint_message}")
                log_msg("🔄 嘗試切換 Key...")
                
                if rotate_key():
                    continue 
                else:
                    raise RuntimeError("All API Keys exhausted")
            
            handle_fatal_error(err_msg)

        except Exception as e:
            log_msg(f"🚨 未預期錯誤: {e}")
            raise e