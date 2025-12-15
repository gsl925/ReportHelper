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
from services_vlm import OllamaVLMService # 匯入新的 VLM 服務

class OllamaManager:
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
    config.load_settings()
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
            'ollama_vlm': OllamaVLMService(api_url=config.OLLAMA_API_URL, model=config.OLLAMA_VLM_MODEL), # 初始化 VLM 服務
            'pptx': PptxService(),
            'pptx_filename': config.MASTER_PPTX_FILENAME
        }

        root = ThemedTkinterDnD(themename="litera")
        
        controller = AppController(None, services, prompts, base_path, ollama_manager)
        ui = AppUI(root, controller)
        controller.ui = ui
        controller.start_background_tasks()

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
