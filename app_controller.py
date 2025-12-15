# app_controller.py (實現附加邏輯)

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

    def start_background_tasks(self):
        if self.ui is None: return
        self.ui.update_status("正在啟動 Ollama 服務，請稍候...", "info")
        threading.Thread(target=self._ollama_status_worker, daemon=True).start()

    def _ollama_status_worker(self):
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
        self.ui.update_status("Ollama 已就緒！請選擇操作。", "success")
        self.ui.set_all_buttons_state("normal")

    def _on_ollama_failed(self):
        self.ui.update_status("Ollama 啟動失敗！", "danger")

    def handle_drop(self, event):
        self.ui.on_drag_leave(event)
        filepaths = self.ui.root.tk.splitlist(event.data)
        if not filepaths: return
        first_file_ext = os.path.splitext(filepaths[0])[1].lower()
        if first_file_ext in ['.png', '.jpg', '.jpeg']:
            if len(filepaths) > 1:
                self.show_warning("多檔案限制", "VLM 模式一次只能處理一張圖片。\n將只分析第一張圖片。")
            image_path = filepaths[0]
            try:
                image = Image.open(image_path)
                self._start_vlm_worker(image)
            except Exception as e:
                self.show_error("讀取圖片檔案失敗", e)
        else:
            self.process_file_list(filepaths)

    def handle_upload_or_paste(self):
        filetypes = (("文字與郵件", "*.txt *.msg"), ("所有檔案", "*.*"))
        filepaths = filedialog.askopenfilenames(filetypes=filetypes)
        if filepaths: self.process_file_list(filepaths)

    def process_file_list(self, filepaths):
        self.ui.set_input_text("", append=False)
        self.ui.set_genai_output_text("")
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
        self.ui.update_status("正在呼叫 LLM 進行分析...", "info")
        self.ui.set_all_buttons_state("disabled")
        threading.Thread(target=self._ollama_worker, args=(full_prompt,), daemon=True).start()

    def _ollama_worker(self, prompt):
        try:
            generated_text = self.services['ollama'].generate(prompt)
            self.ui.root.after(0, self.ui.set_genai_output_text, generated_text)
            self.ui.root.after(0, self.ui.update_status, "LLM 分析報告生成成功！", "success")
        except Exception as e:
            self.ui.root.after(0, self.show_error, "LLM 生成失敗", e)
        finally:
            self.ui.root.after(0, self.ui.set_all_buttons_state, "normal")

    def handle_ppt_generation(self):
        genai_text = self.ui.get_genai_output()
        project_name = self.ui.get_project_name()
        pptx_path = os.path.join(self.base_path, self.services['pptx_filename'])
        try:
            count = self.services['pptx'].add_to_presentation(pptx_path, genai_text, project_name)
            messagebox.showinfo("新增成功", f"成功將 {count} 張投影片新增至\n'{os.path.basename(pptx_path)}'！")
        except PermissionError: self.show_error("權限錯誤", f"無法儲存檔案 '{os.path.basename(pptx_path)}'。\n請先將該 PowerPoint 檔案關閉！")
        except Exception as e: self.show_error("生成 PPT 失敗", e)

    def handle_vlm_generation_from_clipboard(self):
        try:
            image = ImageGrab.grabclipboard()
            if not isinstance(image, Image.Image):
                self.show_warning("無圖片", "剪貼簿中沒有找到圖片。")
                return
            self._start_vlm_worker(image)
        except Exception as e:
            self.show_error("讀取剪貼簿失敗", e)

    def handle_vlm_generation_from_file(self):
        filetypes = (("圖片檔案", "*.png *.jpg *.jpeg"), ("所有檔案", "*.*"))
        filepath = filedialog.askopenfilename(filetypes=filetypes)
        if not filepath: return
        try:
            image = Image.open(filepath)
            self._start_vlm_worker(image)
        except Exception as e:
            self.show_error("讀取圖片檔案失敗", e)

    def _start_vlm_worker(self, image_obj):
        self.ui.update_status("正在呼叫 VLM 進行文字辨識...", "info")
        self.ui.set_all_buttons_state("disabled")
        threading.Thread(target=self._vlm_worker, args=(image_obj,), daemon=True).start()

    # --- 這是本次修改的核心 ---
    def _vlm_worker(self, image_obj):
        """VLM 背景工作函式 (僅執行 OCR)，並根據 UI 選擇附加或覆蓋。"""
        try:
            extracted_text = self.services['ollama_vlm'].get_text_from_image(image_obj)
            
            is_appending = self.ui.is_append_mode()
            
            # 如果是附加模式，且左側已有內容，則加上分隔線
            if is_appending and self.ui.get_input_text():
                text_to_insert = f"\n\n---\n[下一張圖片內容]\n---\n\n{extracted_text}"
            else:
                text_to_insert = extracted_text

            # 如果不是附加模式 (即開始新任務)，則清空右側的舊結果
            if not is_appending:
                self.ui.root.after(0, self.ui.set_genai_output_text, "")

            # 更新左側文字區 (附加或覆蓋)
            self.ui.root.after(0, self.ui.set_input_text, text_to_insert, is_appending)
            self.ui.root.after(0, self.ui.update_status, "VLM 辨識完成！可繼續附加或點擊步驟 B。", "success")

        except Exception as e:
            self.ui.root.after(0, self.show_error, "VLM 辨識失敗", e)
        finally:
            self.ui.root.after(0, self.ui.set_all_buttons_state, "normal")

    def show_error(self, title, message):
        messagebox.showerror(title, str(message))
        self.ui.update_status(f"錯誤: {title}", "danger")

    def show_warning(self, title, message):
        messagebox.showwarning(title, message)
        self.ui.update_status(f"警告: {title}", "warning")
