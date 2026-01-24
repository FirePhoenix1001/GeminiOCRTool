import os
import subprocess

TARGET_KEY = "GEMINI_API_KEY"

def check_registry():
    """ 檢查 Windows 登錄檔中是否存在該變數 """
    # 指令意思：去 HKCU\Environment (當前使用者的環境變數) 查詢 GEMINI_API_KEY
    command = f'REG QUERY "HKCU\\Environment" /V {TARGET_KEY}'
    
    # capture_output=True 會把結果抓回來，不會直接印在螢幕上
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    
    # returncode 為 0 代表找到了(存在)，1 代表找不到(不存在)
    return result.returncode == 0

print(f"\n{'='*40}")
print(f"🧪 實驗開始：刪除 {TARGET_KEY}")
print(f"{'='*40}\n")

# --- 步驟 1: 從 Python 當前記憶體刪除 ---
print("[步驟 1] 清除 Python 執行環境 (os.environ)...")
if TARGET_KEY in os.environ:
    del os.environ[TARGET_KEY]
    print("  -> ✅ 已從 Python 記憶體移除。")
else:
    print("  -> ℹ️ Python 記憶體中原本就不存在。")

# --- 步驟 2: 從 Windows 系統刪除 (登錄檔) ---
print("\n[步驟 2] 清除 Windows 系統設定 (Registry)...")
# /F = 強制刪除, /V = 指定變數名稱
cmd = f'REG DELETE "HKCU\\Environment" /V {TARGET_KEY} /F'
result = subprocess.run(cmd, shell=True, capture_output=True, text=True)

if result.returncode == 0:
    print("  -> ✅ 系統刪除指令執行成功。")
elif "找不到" in result.stderr or "ERROR: The system was unable to find" in result.stderr:
    print("  -> ℹ️ 系統中原本就不存在 (找不到該登錄機碼)。")
else:
    print(f"  -> ⚠️ 指令執行回傳非預期結果: {result.stderr.strip()}")


print(f"\n{'='*40}")
print("🔍 驗證時刻：檢查它是否還活著？")
print(f"{'='*40}")

# --- 驗證 A: Python 內 ---
python_exist = os.environ.get(TARGET_KEY)
if python_exist:
    print(f"❌ [失敗] Python 內仍然看得到: {python_exist}")
else:
    print("✅ [通過] Python 內已讀取不到 (None)")

# --- 驗證 B: Windows 系統內 ---
registry_exist = check_registry()
if registry_exist:
    print("❌ [失敗] Windows 登錄檔內仍然存在！(重開機後會復活)")
else:
    print("✅ [通過] Windows 登錄檔內已消失！(完全刪除)")

print("\n實驗結束。")