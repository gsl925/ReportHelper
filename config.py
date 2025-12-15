# config.py (修正並重構設定載入邏輯)

import json
import os
import sys

# ==============================================================================
# 1. 可由 settings.json 覆寫的全域變數 (提供預設值)
# ==============================================================================
OLLAMA_API_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "deepseek-r1:14b"
OLLAMA_VLM_MODEL = "deepseek-ocr:3b"
MASTER_PPTX_FILENAME = "Weekly Report_JimChuang.pptx"
JPEG_QUALITY = 80  # 預設值 85，在速度和品質之間取得良好平衡

# ==============================================================================
# 2. 開發者定義的半固定設定 (不放入 settings.json)
# ==============================================================================
PROMPT_SINGLE_FILE = 'prompt_single_issue.txt'
PROMPT_MULTI_FILE = 'prompt_multi_issue.txt'
OCR_LANGUAGES = 'chi_tra+chi_sim+eng'
STAR_KEYWORDS = ("情境", "任務", "行動", "結果", "Situation", "Task", "Action", "Result")
SETTINGS_FILENAME = "settings.json"

# ==============================================================================
# 3. 設定載入邏輯 (全新重構)
# ==============================================================================
def get_base_path():
    """獲取程式的基礎路徑，支援打包後的 .exe"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(__file__)

def load_settings():
    """
    載入外部 settings.json 檔案，並智慧地創建或更新它。
    """
    global OLLAMA_API_URL, OLLAMA_MODEL, OLLAMA_VLM_MODEL, MASTER_PPTX_FILENAME, JPEG_QUALITY

    base_path = get_base_path()
    settings_path = os.path.join(base_path, SETTINGS_FILENAME)

    # 步驟 1: 定義程式碼中的預設值
    default_settings = {
        "OLLAMA_API_URL": OLLAMA_API_URL,
        "OLLAMA_MODEL": OLLAMA_MODEL,
        "OLLAMA_VLM_MODEL": OLLAMA_VLM_MODEL,
        "MASTER_PPTX_FILENAME": MASTER_PPTX_FILENAME,
        "JPEG_QUALITY": JPEG_QUALITY
    }

    # 步驟 2: 嘗試載入使用者設定檔
    user_settings = {}
    file_exists = os.path.exists(settings_path)
    if file_exists:
        try:
            with open(settings_path, 'r', encoding='utf-8') as f:
                user_settings = json.load(f)
        except (json.JSONDecodeError, TypeError):
            print(f"警告: {SETTINGS_FILENAME} 檔案已損壞，將使用預設值並覆蓋。")
            file_exists = False # 視為檔案不存在，以便觸發回寫

    # 步驟 3: 合併設定 (使用者的設定會覆蓋預設設定)
    final_settings = default_settings.copy()
    final_settings.update(user_settings)

    # 步驟 4: 用最終的設定來更新全域變數
    try:
        OLLAMA_API_URL = str(final_settings["OLLAMA_API_URL"])
        OLLAMA_MODEL = str(final_settings["OLLAMA_MODEL"])
        OLLAMA_VLM_MODEL = str(final_settings["OLLAMA_VLM_MODEL"])
        MASTER_PPTX_FILENAME = str(final_settings["MASTER_PPTX_FILENAME"])
        JPEG_QUALITY = int(final_settings["JPEG_QUALITY"])
    except (KeyError, ValueError) as e:
        print(f"警告: 設定檔中存在無效的值，將使用預設值。錯誤: {e}")
        # 如果值有問題，回退到預設值
        OLLAMA_API_URL, OLLAMA_MODEL, OLLAMA_VLM_MODEL, MASTER_PPTX_FILENAME, JPEG_QUALITY = default_settings.values()


    # 步驟 5: 如果需要，回寫更新後的設定檔
    # (如果檔案不存在、已損壞，或合併後的設定與原始使用者設定不同)
    if not file_exists or final_settings != user_settings:
        if file_exists:
            print(f"正在更新 {SETTINGS_FILENAME}，加入新的或遺失的設定選項...")
        else:
            print(f"未找到 {SETTINGS_FILENAME}，正在創建預設設定檔...")
        
        try:
            with open(settings_path, 'w', encoding='utf-8') as f:
                json.dump(final_settings, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"錯誤: 無法寫入設定檔 {settings_path}。詳情: {e}")

