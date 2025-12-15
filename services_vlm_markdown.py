# services_vlm.py (升級為 Markdown 輸出)

import base64
import requests
import json
from PIL import Image
import io
import re

class OllamaVLMService:
    """與支援視覺的 Ollama 模型進行通訊"""

    def __init__(self, api_url, model):
        self.api_url = api_url
        self.model = model

    def _preprocess_and_encode_image(self, image_obj: Image.Image, max_size_px=1920):
        """對圖片進行預處理（調整尺寸）並編碼為 Base64。"""
        w, h = image_obj.size
        if w > max_size_px or h > max_size_px:
            if w > h:
                new_w = max_size_px
                new_h = int(h * (max_size_px / w))
            else:
                new_h = max_size_px
                new_w = int(w * (max_size_px / h))
            print(f"圖片尺寸過大 ({w}x{h})，已縮小至 {new_w}x{new_h} 進行處理。")
            image_obj = image_obj.resize((new_w, new_h), Image.Resampling.LANCZOS)

        buffered = io.BytesIO()
        if image_obj.mode in ('RGBA', 'P'):
            image_obj = image_obj.convert('RGB')
        image_obj.save(buffered, format="JPEG", quality=90)
        return base64.b64encode(buffered.getvalue()).decode('utf-8')

    # --- 這是本次修改的核心 ---
    def get_text_from_image(self, image_obj: Image.Image):
        """
        使用 VLM 從圖片中提取結構化的 Markdown 文字。
        """
        try:
            base64_image = self._preprocess_and_encode_image(image_obj)

            # --- 全新的 Markdown Prompt ---
            prompt_text = """
You are an intelligent OCR tool. Your task is to analyze the provided image and transcribe its content into a well-structured Markdown format.

**Instructions:**
1.  Identify structural elements like headings, lists, bold/italic text, and code blocks.
2.  Use appropriate Markdown syntax (e.g., `#` for headings, `*` or `-` for list items).
3.  Preserve the original line breaks and paragraph structure as much as possible.
4.  Your output should ONLY be the Markdown content. Do not add any explanations or introductory sentences.
"""

            payload = {
                "model": self.model,
                "prompt": prompt_text,
                "images": [base64_image],
                "stream": False,
                "options": {
                    "temperature": 0.0
                }
            }

            response = requests.post(self.api_url, json=payload, timeout=600)
            response.raise_for_status()
            
            response_data = response.json()
            generated_text = response_data.get('response', '').strip()

            # --- 新增：清理 Markdown 程式碼區塊標籤 ---
            # 模型有時會習慣性地用 ```markdown ... ``` 包裹輸出
            # 我們用正則表達式來提取中間的內容
            match = re.search(r'```(markdown)?\s*(.*?)\s*```', generated_text, re.DOTALL)
            if match:
                return match.group(2).strip()
            else:
                return generated_text

        except requests.exceptions.Timeout:
            raise Exception("請求 VLM API 超時。\n可能是圖片太複雜或電腦效能不足。\n請嘗試使用尺寸更小的圖片。")
        except Exception as e:
            raise Exception(f"處理 VLM 回應時發生未知錯誤：{e}")

    # (analyze_image_with_star 函式保持不變，作為備用)
    def analyze_image_with_star(self, image_obj: Image.Image, project_name: str):
        pass
