# main.py
# 主程式入口 (v4.1 - 初始化順序修正)

import os
import sys
import time
import subprocess
import requests
import ttkbootstrap as ttk
from tkinter import messagebox
from tkinterdnd2 import TkinterDnD

import config
from app_ui import AppUI
from app_controller import AppController
from services import FileProcessorService, OllamaService, PptxService

class OllamaManager:
    # ... OllamaManager 類別的程式碼保持不變，此處省略以保持簡潔 ...
    def __init__(self, api_base_url="http://localhost:11434"):
        self.api_base_url = api_base_url
        self.ollama_process = None
        self.started_by_app = False

    def _is_server_running(self):
        try:
            requests.head(self.api_base_url, timeout=3)
            return True
        except (requests.exceptions.ConnectionError, requests.exceptions.ReadTimeout):
            return False

    def start_server_non_blocking(self):
        if self._is_server_running():
            print("偵測到 Ollama 服務已在運行。")
            self.started_by_app = False
            return True

        print("Ollama 服務未運行，正在嘗試在背景自動啟動...")
        try:
            creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            self.ollama_process = subprocess.Popen(["ollama", "serve"], creationflags=creationflags)
            self.started_by_app = True
            print(f"Ollama 服務已啟動，主程序 ID: {self.ollama_process.pid}")
            return True
        except FileNotFoundError:
            messagebox.showerror("錯誤", "找不到 'ollama' 指令。\n請確認您已正確安裝 Ollama 且其路徑已加入系統環境變數。")
            return False
        except Exception as e:
            messagebox.showerror("啟動失敗", f"自動啟動 Ollama 服務時發生錯誤：\n{e}")
            return False

    def stop_server(self):
        if not self.ollama_process or not self.started_by_app: return

        print(f"正在關閉由本程式啟動的 Ollama 服務 (主程序 ID: {self.ollama_process.pid})...")
        try:
            if sys.platform == "win32":
                command = f"taskkill /F /PID {self.ollama_process.pid} /T"
                subprocess.run(command, capture_output=True, check=False)
                print("Ollama 服務關閉指令已發送。")
            else:
                self.ollama_process.terminate()
                self.ollama_process.wait(timeout=5)
                print("Ollama 服務已成功關閉。")
        except Exception as e:
            print(f"關閉 Ollama 服務時發生錯誤: {e}")
        finally:
            self.ollama_process = None
            self.started_by_app = False

class ThemedTkinterDnD(TkinterDnD.Tk):
    def __init__(self, *args, **kwargs):
        themename = kwargs.pop('themename', 'litera')
        super().__init__(*args, **kwargs)
        ttk.Style(theme=themename)

def main():
    # --- 關鍵修改：在所有操作之前載入設定 ---
    config.load_settings()
    # --- 修改結束 ---
    ollama_manager = OllamaManager()

    if getattr(sys, 'frozen', False): base_path = os.path.dirname(sys.executable)
    else: base_path = os.path.dirname(__file__)

    try:
        with open(os.path.join(base_path, config.PROMPT_SINGLE_FILE), 'r', encoding='utf-8') as f: single_prompt = f.read()
        with open(os.path.join(base_path, config.PROMPT_MULTI_FILE), 'r', encoding='utf-8') as f: multi_prompt = f.read()
        prompts = {'single': single_prompt, 'multi': multi_prompt}
    except FileNotFoundError as e:
        messagebox.showerror("錯誤", f"找不到必要的 Prompt 檔案: {e.filename}")
        return
    
    try:
        services = {
            'file_processor': FileProcessorService(),
            'ollama': OllamaService(api_url=config.OLLAMA_API_URL, model=config.OLLAMA_MODEL),
            'pptx': PptxService(),
            'pptx_filename': config.MASTER_PPTX_FILENAME
        }

        root = ThemedTkinterDnD(themename="litera")
        
        # --- 修正後的初始化順序 ---
        # 1. 先建立 Controller，但此時不傳入 ui
        controller = AppController(None, services, prompts, base_path, ollama_manager)
        # 2. 建立 UI，並將 controller 傳給它
        ui = AppUI(root, controller)
        # 3. 將建立好的 UI 實例賦值給 controller
        controller.ui = ui
        # 4. 現在 controller.ui 已經有值了，可以安全地呼叫啟動方法
        controller.start_background_tasks()
        # --- 修正結束 ---

        def on_closing():
            if messagebox.askokcancel("退出", "您確定要退出報告整理小幫手嗎？"):
                print("正在準備退出...")
                ollama_manager.stop_server()
                root.destroy()

        root.protocol("WM_DELETE_WINDOW", on_closing)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("應用程式啟動失敗", f"發生嚴重錯誤：\n{e}")
    finally:
        ollama_manager.stop_server()

if __name__ == "__main__":
    main()
