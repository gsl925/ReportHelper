# app_ui.py

import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.widgets.scrolled import ScrolledText
from ttkbootstrap.constants import *
from tkinterdnd2 import DND_FILES

class AppUI:
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.root.title("報告整理小幫手 v16.1 (智慧啟動版)")
        self.root.geometry("1200x750")

        paned_window = ttk.Panedwindow(root, orient=HORIZONTAL)
        paned_window.pack(fill=BOTH, expand=True, padx=10, pady=10)

        self.left_frame = tk.Frame(paned_window) 
        style = ttk.Style.get_instance()
        bg_color = style.colors.get('bg')
        self.left_frame.config(background=bg_color, padx=10, pady=10)

        self.create_left_panel(self.left_frame)
        paned_window.add(self.left_frame, weight=1)

        right_frame = ttk.Frame(paned_window, padding=10)
        self.create_right_panel(right_frame)
        paned_window.add(right_frame, weight=1)

        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.controller.handle_drop)
        self.original_bg = str(self.left_frame.cget("background"))
        self.root.dnd_bind('<<DragEnter>>', self.on_drag_enter)
        self.root.dnd_bind('<<DragLeave>>', self.on_drag_leave)

    def create_left_panel(self, parent):
        project_frame = ttk.Frame(parent)
        project_frame.pack(fill=X, pady=(0, 15))
        ttk.Label(project_frame, text="專案名稱 (選填):").pack(side=LEFT, padx=(0, 10))
        self.project_name_entry = ttk.Entry(project_frame)
        self.project_name_entry.pack(side=LEFT, fill=X, expand=True)

        upload_paste_frame = ttk.Frame(parent)
        upload_paste_frame.pack(fill=X, pady=(0, 10))
        self.upload_button = ttk.Button(upload_paste_frame, text="1. 上傳/貼上", command=self.controller.handle_upload_or_paste, bootstyle="info")
        self.upload_button.pack(side=LEFT, padx=(0, 5))
        self.status_label = ttk.Label(upload_paste_frame, text="請上傳、貼上或拖曳檔案至此...", bootstyle="primary")
        self.status_label.pack(side=LEFT, padx=10)

        text_frame = ttk.Labelframe(parent, text="步驟 A: 辨識結果 (原始文字)", padding=5)
        text_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        self.text_area = ScrolledText(text_frame, wrap=WORD, autohide=True)
        self.text_area.pack(fill=BOTH, expand=True)
        
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=X, ipady=5)
        
        # --- 變更：預設禁用按鈕 ---
        self.single_button = ttk.Button(button_frame, text="步驟 B: 分析為「單一問題」報告", command=lambda: self.controller.handle_ollama_generation('single'), bootstyle="primary", state="disabled")
        self.single_button.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        
        self.multi_button = ttk.Button(button_frame, text="步驟 B: 分析為「多個問題」報告", command=lambda: self.controller.handle_ollama_generation('multi'), bootstyle="success", state="disabled")
        self.multi_button.pack(side=LEFT, fill=X, expand=True, padx=(5, 0))

    def create_right_panel(self, parent):
        genai_frame = ttk.Labelframe(parent, text="步驟 C: Ollama 生成結果", padding=5)
        genai_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        self.genai_output_area = ScrolledText(genai_frame, wrap=WORD, autohide=True)
        self.genai_output_area.pack(fill=BOTH, expand=True)

        self.ppt_button = ttk.Button(parent, text="步驟 D: 新增至彙總簡報", command=self.controller.handle_ppt_generation, bootstyle="danger")
        self.ppt_button.pack(fill=X, ipady=5)

    def on_drag_enter(self, event):
        self.left_frame.config(background="#e0e0e0")
        return event.action

    def on_drag_leave(self, event):
        self.left_frame.config(background=self.original_bg)

    def get_project_name(self): return self.project_name_entry.get().strip()
    def get_input_text(self): return self.text_area.get("1.0", tk.END).strip()
    def get_genai_output(self): return self.genai_output_area.get("1.0", tk.END).strip()

    def set_input_text(self, text, append=False):
        if not append: self.text_area.delete('1.0', tk.END)
        self.text_area.insert(tk.END, text)

    def set_genai_output_text(self, text):
        self.genai_output_area.delete('1.0', tk.END)
        self.genai_output_area.insert(tk.END, text)

    def update_status(self, text, style="primary"):
        self.status_label.config(text=text, bootstyle=style)
        self.root.update_idletasks()
        
    # --- 新增：專門控制分析按鈕的方法 ---
    def set_generator_buttons_state(self, state):
        self.single_button.config(state=state)
        self.multi_button.config(state=state)
