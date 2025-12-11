# config.py
# 設定載入器，負責讀取外部 settings.json，並定義半固定設定

import json
import os
import sys

# ==============================================================================
# 1. 可由 settings.json 覆寫的全域變數 (提供預設值)
# ==============================================================================
OLLAMA_API_URL = "http://localhost:11434/api/generate"  # Ollama 的 API 位址
OLLAMA_MODEL = "deepseek-r1:14b"  # 您想要使用的模型，例如 "llama3", "mistral" 等
MASTER_PPTX_FILENAME = "Weekly Report_JimChuang.pptx"


# ==============================================================================
# 2. 開發者定義的半固定設定 (不放入 settings.json)
# ==============================================================================
# Prompt 檔案的名稱是固定的，程式會預期在主目錄下找到它們
PROMPT_SINGLE_FILE = 'prompt_single_issue.txt'
PROMPT_MULTI_FILE = 'prompt_multi_issue.txt'

# OCR 相關設定
OCR_LANGUAGES = 'chi_tra+chi_sim+eng'

# PowerPoint 排版相關的關鍵字
STAR_KEYWORDS = ("情境", "任務", "行動", "結果", "Situation", "Task", "Action", "Result")

# 外部設定檔的固定名稱
SETTINGS_FILENAME = "settings.json"


# ==============================================================================
# 3. 設定載入邏輯
# ==============================================================================

def get_base_path():
    """獲取程式的基礎路徑，支援打包後的 .exe"""
    if getattr(sys, 'frozen', False):
        # 如果程式被打包成 .exe
        return os.path.dirname(sys.executable)
    else:
        # 如果是作為 .py 腳本運行
        return os.path.dirname(__file__)

def load_settings():
    """
    載入外部 settings.json 檔案。
    如果檔案不存在，則使用預設值並創建一個新的設定檔。
    這個函數只會修改 OLLAMA_MODEL, MASTER_PPTX_FILENAME, OLLAMA_API_URL。
    """
    global OLLAMA_MODEL, MASTER_PPTX_FILENAME, OLLAMA_API_URL

    base_path = get_base_path()
    settings_path = os.path.join(base_path, SETTINGS_FILENAME)

    # 定義哪些設定是使用者可以修改的
    defaults = {
        "OLLAMA_MODEL": OLLAMA_MODEL,
        "MASTER_PPTX_FILENAME": MASTER_PPTX_FILENAME,
        "OLLAMA_API_URL": OLLAMA_API_URL
    }

    if os.path.exists(settings_path):
        # 如果檔案存在，讀取它
        print(f"正在從 {SETTINGS_FILENAME} 載入設定...")
        try:
            with open(settings_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            
            # 用讀取到的值更新全域變數，如果某個鍵不存在，則使用預設值
            OLLAMA_MODEL = settings.get("OLLAMA_MODEL", defaults["OLLAMA_MODEL"])
            MASTER_PPTX_FILENAME = settings.get("MASTER_PPTX_FILENAME", defaults["MASTER_PPTX_FILENAME"])
            OLLAMA_API_URL = settings.get("OLLAMA_API_URL", defaults["OLLAMA_API_URL"])
            print("設定載入成功！")
            
            # 檢查是否有新的預設鍵，並更新檔案，以便使用者知道有新選項
            _update_settings_file_if_needed(settings_path, defaults, settings)

        except (json.JSONDecodeError, TypeError) as e:
            print(f"錯誤：{SETTINGS_FILENAME} 格式不正確，將使用預設設定。錯誤詳情: {e}")
            _create_default_settings(settings_path, defaults)
    else:
        # 如果檔案不存在，創建一個
        print(f"未找到 {SETTINGS_FILENAME}，正在創建預設設定檔...")
        _create_default_settings(settings_path, defaults)

def _create_default_settings(path, defaults):
    """創建一個預設的設定檔"""
    try:
        with open(path, 'w', encoding='utf-8') as f:
            # 使用 indent=4 讓 JSON 檔案格式更美觀，易於手動編輯
            json.dump(defaults, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"錯誤：無法創建預設設定檔 {path}。錯誤詳情: {e}")

def _update_settings_file_if_needed(path, defaults, current_settings):
    """如果程式更新後多了新的可配置項，自動加到現有設定檔中"""
    needs_update = False
    for key, value in defaults.items():
        if key not in current_settings:
            current_settings[key] = value
            needs_update = True
    
    if needs_update:
        print("偵測到新的設定選項，正在更新 settings.json...")
        _create_default_settings(path, current_settings)

