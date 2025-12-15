# services_vlm.py (使用可配置的壓縮品質)

import base64
import requests
import json
from PIL import Image
import io
import re
import config # 匯入 config 模組

class OllamaVLMService:
    def __init__(self, api_url, model):
        self.api_url = api_url
        self.model = model

    def _preprocess_and_encode_image(self, image_obj: Image.Image, max_size_px=1920):
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
        
        # --- 這是本次修改的核心 ---
        # 使用從 config 模組讀取到的 JPEG_QUALITY 值
        print(f"使用 JPEG 品質 {config.JPEG_QUALITY} 進行壓縮。")
        image_obj.save(buffered, format="JPEG", quality=config.JPEG_QUALITY)
        # --- 修改結束 ---
        
        return base64.b64encode(buffered.getvalue()).decode('utf-8')

    def get_text_from_image(self, image_obj: Image.Image):
        try:
            base64_image = self._preprocess_and_encode_image(image_obj)
            
            prompt_text = "You are an OCR tool. Your only task is to transcribe the text from the following image. Output ONLY the raw text content. Do not add any other words, explanations, or markdown formatting."

            payload = {
                "model": self.model,
                "prompt": prompt_text,
                "images": [base64_image],
                "stream": False,
                "options": { "temperature": 0.0 }
            }

            response = requests.post(self.api_url, json=payload, timeout=300)
            response.raise_for_status()
            
            response_data = response.json()
            generated_text = response_data.get('response', '').strip()
            
            match = re.search(r'```(markdown)?\s*(.*?)\s*```', generated_text, re.DOTALL)
            return match.group(2).strip() if match else generated_text

        except requests.exceptions.Timeout:
            raise Exception("請求 VLM API 超時。\n可能是圖片太複雜或電腦效能不足。")
        except Exception as e:
            raise Exception(f"處理 VLM 回應時發生未知錯誤：{e}")
