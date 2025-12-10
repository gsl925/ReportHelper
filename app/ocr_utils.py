# app/ocr_utils.py
import io
import os
from PIL import Image, ImageOps
import numpy as np
import cv2
from paddleocr import PaddleOCR
import pytesseract
from pdf2image import convert_from_bytes

# 初始化 PaddleOCR（語言可調）
OCR = PaddleOCR(use_angle_cls=True, lang='ch')  # 可設為 'ch', 'en', 'ch_en'


def image_preprocess_pil(pil_img: Image.Image, enlarge: bool = True, debug_dir: str = None) -> Image.Image: 
    """ 預處理圖片：灰階 -> CLAHE 對比增強 -> 二值化 -> 去噪 -> deskew -> 放大 將每個步驟的圖片存檔到 debug_dir 方便排查 """ 
    os.makedirs(debug_dir, exist_ok=True) if debug_dir else None
    img = pil_img.convert('RGB')
    if debug_dir:
        img.save(os.path.join(debug_dir, 'original.png'))


    arr = np.array(img)
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    if debug_dir:
        Image.fromarray(gray).save(os.path.join(debug_dir, 'gray.png'))


    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
    cl = clahe.apply(gray)
    if debug_dir:
        Image.fromarray(cl).save(os.path.join(debug_dir, 'clahe.png'))


    th = cv2.adaptiveThreshold(cl, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    if debug_dir:
        Image.fromarray(th).save(os.path.join(debug_dir, 'threshold.png'))


    deno = cv2.medianBlur(th, 3)
    if debug_dir:
        Image.fromarray(deno).save(os.path.join(debug_dir, 'denoise.png'))


    coords = np.column_stack(np.where(deno > 0))
    angle = 0.0
    if coords.shape[0] > 0:
        rect = cv2.minAreaRect(coords)
        angle = rect[-1]
        if angle < -45:
            angle = -(90 + angle)
        else:
            angle = -angle


    (h, w) = deno.shape
    center = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(deno, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    if debug_dir:
        Image.fromarray(rotated).save(os.path.join(debug_dir, 'deskew.png'))


    pil_out = Image.fromarray(rotated)
    if enlarge:
        pil_out = pil_out.resize(
        (int(pil_out.width * 1.5), int(pil_out.height * 1.5)),
        Image.Resampling.BICUBIC
        )
    if debug_dir:
        pil_out.save(os.path.join(debug_dir, 'resized.png'))


    return pil_out

def ocr_from_image_bytes(bytes_data: bytes, use_paddle: bool = True, tesseract_fallback: bool = False, debug_dir: str = None) -> str: 
    """ 從影像或 PDF bytes 擷取文字 debug_dir 可指定輸出預處理後圖片 """ 
    text_blocks = [] 
    try: 
        img = Image.open(io.BytesIO(bytes_data)) 
    except Exception: 
        pages = convert_from_bytes(bytes_data, dpi=200) 
        if len(pages) == 0: 
            return "" 
        img = pages[0]
    pre = image_preprocess_pil(img, debug_dir=debug_dir)
    if use_paddle:
        try:
            res = OCR.ocr(np.array(pre), cls=True)
            lines = []
            for page in res:
                for line in page:
                    text = line[1][0] if len(line) > 1 else ''
                    lines.append(text)
                    text_blocks.append('\n'.join(lines))
        except Exception as e:
            print('PaddleOCR error:', e)
    if tesseract_fallback and not text_blocks:
        try:
            ttxt = pytesseract.image_to_string(pre, lang='chi_sim+eng')
            text_blocks.append(ttxt)
        except Exception as e:
            print('Tesseract error:', e)
    return '\n'.join(text_blocks)
