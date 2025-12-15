# services_vlm.py (最終優化版)

import base64
import requests
import json
from PIL import Image
import io

class OllamaVLMService:
    """與支援視覺的 Ollama 模型進行通訊"""

    def __init__(self, api_url, model):
        self.api_url = api_url
        self.model = model

    def _preprocess_and_encode_image(self, image_obj: Image.Image, max_size_px=1920):
        """
        對圖片進行預處理（調整尺寸）並編碼為 Base64。
        """
        # --- 新增的圖片尺寸調整邏輯 ---
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
        # --- 調整結束 ---

        buffered = io.BytesIO()
        if image_obj.mode in ('RGBA', 'P'):
            image_obj = image_obj.convert('RGB')
        image_obj.save(buffered, format="JPEG", quality=90) # 使用稍高的品質
        return base64.b64encode(buffered.getvalue()).decode('utf-8')

    def get_text_from_image(self, image_obj: Image.Image):
        """
        僅使用 VLM 從圖片中提取純文字 (OCR 功能)。
        """
        try:
            # 使用新的預處理函式
            base64_image = self._preprocess_and_encode_image(image_obj)

            prompt_text = "You are an OCR tool. Your only task is to transcribe the text from the following image. Output ONLY the raw text content. Do not add any other words, explanations, or markdown formatting."

            payload = {
                "model": self.model,
                "prompt": prompt_text,
                "images": [base64_image],
                "stream": False,
                "options": {
                    "temperature": 0.0
                }
            }

            # 保持 300 秒超時，但現在成功的機率會高很多
            response = requests.post(self.api_url, json=payload, timeout=600)
            response.raise_for_status()
            
            response_data = response.json()
            return response_data.get('response', '').strip()

        except requests.exceptions.Timeout:
            raise Exception("請求 VLM API 超時。\n可能是圖片太複雜或電腦效能不足。\n請嘗試使用尺寸更小的圖片。")
        except requests.exceptions.ConnectionError:
            raise Exception(f"無法連接至 Ollama VLM API ({self.api_url})。\n請確認 Ollama 正在運行且模型 '{self.model}' 已安裝。")
        except requests.exceptions.RequestException as e:
            try:
                error_detail = e.response.json().get('error', str(e))
                if "model" in error_detail and "not found" in error_detail:
                     raise Exception(f"模型 '{self.model}' 未找到。\n請確認您已在 Ollama 中下載此視覺模型。")
            except:
                pass
            raise Exception(f"請求 Ollama VLM API 時發生錯誤：{e}")
        except Exception as e:
            raise Exception(f"處理 VLM 回應時發生未知錯誤：{e}")

    def analyze_image_with_star(self, image_obj: Image.Image, project_name: str):
        pass
