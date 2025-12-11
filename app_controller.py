# app_controller.py

import os
import sys
import threading
import time
from tkinter import filedialog, messagebox
from PIL import Image, ImageGrab

class AppController:
    def __init__(self, ui, services, prompts, base_path, ollama_manager):
        self.ui = ui
        self.services = services
        self.prompts = prompts
        self.base_path = base_path
        self.ollama_manager = ollama_manager

    # --- 新增：將啟動邏輯移到這裡 ---
    def start_background_tasks(self):
        """在 UI 完全初始化後，啟動所有背景任務"""
        if self.ui is None:
            print("錯誤：UI 尚未初始化，無法啟動背景任務。")
            return
            
        self.ui.update_status("正在啟動 Ollama 服務，請稍候...", "info")
        threading.Thread(target=self._ollama_status_worker, daemon=True).start()

    def _ollama_status_worker(self):
        """在背景執行緒中檢查 Ollama 服務狀態"""
        if not self.ollama_manager.start_server_non_blocking():
            self.ui.root.after(0, self._on_ollama_failed)
            return

        if self.ollama_manager.started_by_app:
            print("正在後台等待 Ollama 服務就緒...")
            while not self.ollama_manager._is_server_running():
                time.sleep(2)
            print("後台偵測到 Ollama 服務已就緒！")

        self.ui.root.after(0, self._on_ollama_ready)

    def _on_ollama_ready(self):
        """Ollama 準備就緒後的回呼函數 (在 UI 執行緒中執行)"""
        self.ui.update_status("Ollama 已就緒！請輸入內容。", "success")
        self.ui.set_generator_buttons_state("normal")

    def _on_ollama_failed(self):
        """Ollama 啟動失敗的回呼函數"""
        self.ui.update_status("Ollama 啟動失敗！", "danger")

    # --- 以下其他方法保持不變 ---
    def handle_drop(self, event):
        self.ui.on_drag_leave(event)
        filepaths = self.ui.root.tk.splitlist(event.data)
        if filepaths: self.process_file_list(filepaths)

    def handle_upload_or_paste(self):
        try:
            image = ImageGrab.grabclipboard()
            if isinstance(image, Image.Image):
                self.process_clipboard_image(image)
                return
        except Exception: pass
        
        filetypes = (("支援的檔案", "*.png *.jpg *.jpeg *.txt *.msg"), ("所有檔案", "*.*"))
        filepaths = filedialog.askopenfilenames(filetypes=filetypes)
        if filepaths: self.process_file_list(filepaths)

    def process_clipboard_image(self, image):
        self.ui.update_status("正在辨識剪貼簿圖片...", "info")
        try:
            if self.ui.get_input_text():
                self.ui.set_input_text(f"\n\n{'='*20} 來自剪貼簿的新增圖片 {'='*20}\n\n", append=True)
            
            extracted_text = self.services['file_processor'].process_image_object(image)
            self.ui.set_input_text(extracted_text, append=True)
            self.ui.update_status("剪貼簿圖片辨識完成！", "success")
        except Exception as e: self.show_error("處理剪貼簿圖片失敗", e)

    def process_file_list(self, filepaths):
        self.ui.set_input_text("")
        total_files = len(filepaths)
        for i, file_path in enumerate(filepaths):
            if i > 0: self.ui.set_input_text(f"\n\n{'='*20} 檔案 {i+1} {'='*20}\n\n", append=True)
            self.ui.update_status(f"處理中 {i+1}/{total_files}: {os.path.basename(file_path)}", "info")
            try:
                file_ext = os.path.splitext(file_path)[1].lower()
                content = ""
                if file_ext in ['.png', '.jpg', '.jpeg']: content = self.services['file_processor'].process_image_object(Image.open(file_path))
                elif file_ext == '.txt': content = self.services['file_processor'].process_text_file(file_path)
                elif file_ext == '.msg': content = self.services['file_processor'].process_msg_file(file_path)
                else:
                    self.show_warning("不支援的格式", f"已跳過不支援的檔案格式: {file_ext}")
                    continue
                self.ui.set_input_text(content, append=True)
            except Exception as e: self.show_error(f"處理檔案 {os.path.basename(file_path)} 失敗", e)
        self.ui.update_status(f"全部 {total_files} 個檔案處理完成！", "success")

    def handle_ollama_generation(self, prompt_type):
        input_content = self.ui.get_input_text()
        if not input_content:
            self.show_warning("內容為空", "請先從左側輸入內容後再進行分析。")
            return

        base_prompt = self.prompts[prompt_type]
        project_name = self.ui.get_project_name()
        project_info = f"The user has specified the project name is: '{project_name}'." if project_name else "The user did not specify a project name."
        final_prompt = base_prompt.replace("{PROJECT_NAME_HOLDER}", project_info)
        full_prompt = f"{final_prompt}\n\n{input_content}"
        
        self.ui.update_status("正在呼叫 Ollama 模型生成報告...", "info")
        self.ui.set_generator_buttons_state("disabled")
        threading.Thread(target=self._ollama_worker, args=(full_prompt,)).start()

    def _ollama_worker(self, prompt):
        try:
            generated_text = self.services['ollama'].generate(prompt)
            self.ui.root.after(0, self.ui.set_genai_output_text, generated_text)
            self.ui.root.after(0, self.ui.update_status, "Ollama 報告生成成功！", "success")
        except Exception as e:
            self.ui.root.after(0, self.show_error, "Ollama 生成失敗", e)
        finally:
            self.ui.root.after(0, self.ui.set_generator_buttons_state, "normal")

    def handle_ppt_generation(self):
        genai_text = self.ui.get_genai_output()
        project_name = self.ui.get_project_name()
        pptx_path = os.path.join(self.base_path, self.services['pptx_filename'])
        try:
            count = self.services['pptx'].add_to_presentation(pptx_path, genai_text, project_name)
            messagebox.showinfo("新增成功", f"成功將 {count} 張投影片新增至\n'{os.path.basename(pptx_path)}'！")
        except PermissionError: self.show_error("權限錯誤", f"無法儲存檔案 '{os.path.basename(pptx_path)}'。\n請先將該 PowerPoint 檔案關閉！")
        except Exception as e: self.show_error("生成 PPT 失敗", e)

    def show_error(self, title, message):
        messagebox.showerror(title, message)
        self.ui.update_status(f"錯誤: {title}", "danger")

    def show_warning(self, title, message):
        messagebox.showwarning(title, message)
        self.ui.update_status(f"警告: {title}", "warning")
