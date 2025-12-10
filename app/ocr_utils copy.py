# app/ocr_utils.py
import io
from PIL import Image, ImageOps
import numpy as np
import cv2
from paddleocr import PaddleOCR
import pytesseract
from pdf2image import convert_from_bytes

# 初始化 PaddleOCR（語言可調）
OCR = PaddleOCR(use_angle_cls=True, lang='ch')  # 可設為 'ch', 'en', 'ch_en'


def image_preprocess_pil(pil_img: Image.Image, enlarge: bool = True) -> Image.Image:
    # 轉灰階 -> 自適應對比 -> 去噪 -> deskew
    img = pil_img.convert('RGB')
    arr = np.array(img)
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    # CLAHE
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
    cl = clahe.apply(gray)
    # 二值化
    th = cv2.adaptiveThreshold(cl,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,11,2)
    # denoise
    deno = cv2.medianBlur(th,3)
    # deskew
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
    center = (w//2, h//2)
    M = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(deno, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    pil_out = Image.fromarray(rotated)
    if enlarge:
        pil_out = pil_out.resize((int(pil_out.width*1.5), int(pil_out.height*1.5)), Image.BICUBIC)
    return pil_out


def ocr_from_image_bytes(bytes_data: bytes, use_paddle: bool = True, tesseract_fallback: bool = True) -> str:
    # supports image bytes and pdf bytes (first page)
    text_blocks = []
    try:
        img = Image.open(io.BytesIO(bytes_data))
    except Exception:
        # maybe PDF
        pages = convert_from_bytes(bytes_data, dpi=200)
        if len(pages) == 0:
            return ""
        img = pages[0]
    # preprocess
    pre = image_preprocess_pil(img)
    if use_paddle:
        try:
            res = OCR.ocr(np.array(pre))
            lines = []
            for line in res:
                for item in line:
                    if isinstance(item, list) and len(item) >= 2:
                        lines.append(item[1][0])
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