import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pytesseract
from PIL import Image, ImageGrab
import os
import pyperclip

# --- 設定 Tesseract 路徑 (僅 Windows 需要) ---
# if os.name == 'nt':
#     pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# --- v3.0: 全自動多問題分析 Prompt ---
COMBINED_PROMPT = """
# Role and Goal
You are an expert business analyst and communication consultant. Your primary goal is to analyze a potentially long and unstructured work discussion, identify ALL distinct issues within it, and generate a separate, complete STAR method report for EACH issue identified.

# Task
Your task is a multi-step process:
1.  **Identify & Separate**: First, read the entire input text and identify all the distinct, unrelated problems or topics.
2.  **Iterate and Analyze**: For EACH distinct problem you have identified, perform the following sub-tasks:
    a. Create a concise and factual title that summarizes that specific problem.
    b. Structure the key information for that problem according to the STAR method (Situation, Task, Action, Result).

# Input Format
A single block of raw text that may contain discussions about one or more different problems. The text may be in Traditional Chinese, Simplified Chinese, or English.

# Output Format and Rules
- The entire output MUST be in **Traditional Chinese (繁體中文)**.
- For each report, use a clear separator and title, like `--- 報告 1：[問題標題] ---`.
- Under each separator, provide the full STAR analysis (Situation, Task, Action, Result) for that specific issue.
- **Objectivity is key**: For each report, stick strictly to the facts related to that issue. Do not mix information between different issues.
- **Handle Missing Information**: If a specific part of the STAR method for an issue cannot be determined from the text, explicitly state "資訊不足".
- **Handle Single-Issue Case**: If the input text only contains one single, coherent issue, just produce one report under the `--- 報告 1 ---` header.
- **Professional Tone**: The language must be formal and professional.

# Example
---
[EXAMPLE INPUT TEXT]
Ben: @Ariel，v3.2.1 的 hotfix 上完了，結帳功能恢復了。
Ariel: 太好了！辛苦了。另外，PM 那邊在問，我們上次討論的會員後台報表卡頓的問題，有進展嗎？他們想明天下午同步一下。
Ben: 报表那个我查了，是资料库 query 没走索引，数据量一大就慢。我正在优化 SQL，明天早上应该可以更新。
Lisa: Hi all, 我是新來的供應商窗口 Lisa。想請問一下，上週五採購單 #PO-2024-05-123 的物料，系統顯示已出貨，但我們倉庫這邊還沒收到，能幫忙查一下物流狀態嗎？
Ben: @Lisa Hi Lisa, 歡迎！我幫妳轉給倉管部門的 David，他會幫妳追蹤。
---
[EXAMPLE OUTPUT]
--- 報告 1：v3.2.1 Hotfix 完成與結帳功能恢復 ---
- **情境 (Situation)**
  - 系統 v3.2.1 版本上線後出現的結帳問題，已透過緊急修補 (Hotfix) 處理。
- **任務 (Task)**
  - 向相關人員同步更新事件處理進度。
- **行動 (Action)**
  - 工程師 (Ben) 在完成修補後，於討論串中通知 Ariel。
- **結果 (Result)**
  - 結帳功能已恢復正常，相關人員已知曉此狀態。

--- 報告 2：會員後台報表卡頓問題調查與修復 ---
- **情境 (Situation)**
  - 會員後台的報表功能在數據量大時會發生卡頓。PM 團隊正在關注此問題的進度。
  - 初步調查發現，根本原因是資料庫查詢未正確使用索引。
- **任務 (Task)**
  - 優化 SQL 查詢以解決報表卡頓問題。
- **行動 (Action)**
  - 工程師 (Ben) 正在進行 SQL 優化工作。
- **結果 (Result)**
  - 預計在次日早上完成更新並解決問題。PM 團隊將在次日下午獲得進度同步。

--- 報告 3：採購單 #PO-2024-05-123 物流狀態查詢 ---
- **情境 (Situation)**
  - 新任供應商窗口 (Lisa) 反應，一筆採購單的物料系統顯示已出貨，但倉庫未收到。
- **任務 (Task)**
  - 協助查詢該筆訂單的實際物流狀態。
- **行動 (Action)**
  - Ben 將此問題轉交給負責倉儲管理的 David 進行追蹤。
- **結果 (Result)**
  - 問題已轉交給權責部門處理，等待倉管部門的回覆。
---

Now, analyze the following text and generate all reports in the specified format.
"""

class ReportHelperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("報告整理小幫手 v3.0 (全自動多問題分析)")
        self.root.geometry("800x600")

        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Top controls
        controls_frame = tk.Frame(main_frame)
        controls_frame.pack(fill=tk.X, pady=(0, 10))

        upload_button = tk.Button(controls_frame, text="1. 上傳檔案", command=self.upload_file)
        upload_button.pack(side=tk.LEFT, padx=(0, 5))

        paste_button = tk.Button(controls_frame, text="2. 從剪貼簿貼上圖片", command=self.paste_from_clipboard)
        paste_button.pack(side=tk.LEFT, padx=5)

        self.status_label = tk.Label(controls_frame, text="請上傳檔案或貼上圖片...", fg="blue")
        self.status_label.pack(side=tk.LEFT, padx=10)

        # Text area
        text_frame = tk.LabelFrame(main_frame, text="辨識結果 (可編輯)", padx=5, pady=5)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.text_area = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD)
        self.text_area.pack(fill=tk.BOTH, expand=True)
        
        # Bottom button
        copy_button = tk.Button(main_frame, text="複製 Prompt 與內容至剪貼簿", command=self.copy_for_genai, bg="#4CAF50", fg="white", height=2, font=("Arial", 10, "bold"))
        copy_button.pack(fill=tk.X)

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=(("圖片檔案", "*.png *.jpg *.jpeg"), ("文字檔案", "*.txt"), ("所有檔案", "*.*")))
        if not file_path: return
        file_ext = os.path.splitext(file_path)[1].lower()
        try:
            if file_ext in ['.png', '.jpg', '.jpeg']:
                self.process_image_object(Image.open(file_path))
            elif file_ext in ['.txt']:
                self.process_text_file(file_path)
            else:
                messagebox.showwarning("不支援的格式", f"不支援的檔案格式: {file_ext}")
        except Exception as e:
            self.handle_error(e)

    def paste_from_clipboard(self):
        try:
            image = ImageGrab.grabclipboard()
            if isinstance(image, Image.Image):
                self.process_image_object(image)
            else:
                messagebox.showinfo("提示", "剪貼簿中沒有圖片。")
                self.status_label.config(text="剪貼簿中沒有圖片。", fg="orange")
        except Exception as e:
            self.handle_error(e, "從剪貼簿讀取圖片時發生錯誤")

    def process_image_object(self, image_obj):
        self.status_label.config(text="正在進行 OCR 辨識 (簡/繁/英)...", fg="blue")
        self.root.update_idletasks()
        try:
            lang_models = 'chi_tra+chi_sim+eng'
            extracted_text = pytesseract.image_to_string(image_obj, lang=lang_models)
            self.text_area.delete('1.0', tk.END)
            self.text_area.insert(tk.INSERT, extracted_text)
            self.status_label.config(text="圖片辨識完成！", fg="green")
        except pytesseract.TesseractNotFoundError:
            messagebox.showerror("Tesseract 未找到", "找不到 Tesseract OCR 引擎。請確認已安裝且路徑正確。")
            self.status_label.config(text="Tesseract 未安裝或路徑錯誤！", fg="red")
        except Exception as e:
            raise e

    def process_text_file(self, file_path):
        self.status_label.config(text="正在讀取文字檔...", fg="blue")
        self.root.update_idletasks()
        try:
            with open(file_path, 'r', encoding='utf-8') as f: content = f.read()
        except Exception:
            try:
                with open(file_path, 'r', encoding='gbk') as f: content = f.read()
            except Exception as e2:
                self.handle_error(e2, "讀取文字檔失敗")
                return
        self.text_area.delete('1.0', tk.END)
        self.text_area.insert(tk.INSERT, content)
        self.status_label.config(text="文字檔讀取完成！", fg="green")

    def copy_for_genai(self):
        content = self.text_area.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("內容為空", "辨識結果為空，無法複製。")
            return
        full_prompt = f"{COMBINED_PROMPT}\n\n{content}"
        pyperclip.copy(full_prompt)
        messagebox.showinfo("複製成功", "全功能 Prompt 及內容已複製到剪貼簿！\n請直接貼到 GenAI 應用中。")

    def handle_error(self, e, custom_msg="處理時發生錯誤"):
        messagebox.showerror("錯誤", f"{custom_msg}：\n{e}")
        self.status_label.config(text="處理失敗！", fg="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = ReportHelperApp(root)
    root.mainloop()
