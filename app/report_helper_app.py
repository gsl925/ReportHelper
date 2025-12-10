# app/report_helper_app.py
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

import streamlit as st
from io import BytesIO
from app.ocr_utils import ocr_from_image_bytes
from app.postprocess import clean_text, load_domain_dict, extract_key_sentences, simple_star_from_sentences, apply_domain_corrections
from app.pptx_export import export_to_pptx
import json
import os

st.set_page_config(page_title='報告整理小幫手', layout='wide')

st.title('報告整理小幫手 (OCR -> STAR -> PPTX)')

st.sidebar.header('設定')
use_paddle = st.sidebar.checkbox('使用 PaddleOCR', value=True)
use_tesseract = st.sidebar.checkbox('Tesseract 備援', value=True)

uploaded = st.file_uploader('上傳影像或 PDF（png/jpg/pdf）', type=['png','jpg','jpeg','pdf'])

# load domain dict
DOMAIN_DICT = load_domain_dict(os.path.join(os.getcwd(), 'domain_dict.json'))

if uploaded is not None:
    b = uploaded.read()
    st.subheader('上傳檔案預覽')
    try:
        st.image(b)
    except Exception:
        st.write('PDF / binary file uploaded')

    with st.spinner('進行 OCR...'):
        raw_text = ocr_from_image_bytes(b, use_paddle=use_paddle, tesseract_fallback=use_tesseract)
    if not raw_text.strip():
        st.error('OCR 未擷取到文字，請嘗試調整上傳圖片或使用更高解析度掃描。')
    raw_text = clean_text(raw_text)
    st.subheader('OCR 文字（請檢查並修正）')
    txt = st.text_area('OCR 原文', value=raw_text, height=300)

    st.subheader('自動抽取關鍵句')
    key_sents = extract_key_sentences(txt, DOMAIN_DICT, topn=8)
    for i, s in enumerate(key_sents):
        st.write(f'{i+1}. {s}')

    st.subheader('產生 STAR (簡易規則)')
    star = simple_star_from_sentences(key_sents)
    st.json(star)

    st.subheader('編輯 STAR（若需）')
    # allow user to edit each section
    s_situation = st.text_area('Situation (每行一個要點)', value='\n'.join(star.get('situation', [])), height=120)
    s_task = st.text_area('Task', value='\n'.join(star.get('task', [])), height=100)
    s_action = st.text_area('Action', value='\n'.join(star.get('action', [])), height=140)
    s_result = st.text_area('Result', value='\n'.join(star.get('result', [])), height=100)

    st.subheader('標題建議')
    # naive title candidates from domain
    t1 = DOMAIN_DICT.get('product_codes', [])[0] + ' 測試異常' if DOMAIN_DICT.get('product_codes') else '測試異常'
    t2 = '當日良率下降與處置'
    t3 = '需3日追蹤之臨時處置'
    titles = [t1, t2, t3]
    st.write('候選標題：')
    for t in titles:
        st.write('- ', t)

    st.subheader('輸出與匯出')
    chosen_title = st.text_input('最終標題（可修改）', value=titles[0])
    if st.button('匯出 PPTX'):
        star_final = {
            'situation': [l for l in s_situation.split('\n') if l.strip()],
            'task': [l for l in s_task.split('\n') if l.strip()],
            'action': [l for l in s_action.split('\n') if l.strip()],
            'result': [l for l in s_result.split('\n') if l.strip()]
        }
        outpath = export_to_pptx('report_star.pptx', chosen_title, star_final)
        with open(outpath, 'rb') as f:
            st.download_button('下載 PPTX', data=f, file_name='report_star.pptx', mime='application/vnd.openxmlformats-officedocument.presentationml.presentation')

    st.info('提示：請務必檢查並修正 OCR 結果以避免 LLM 或報表產生錯誤訊息。')
else:
    st.info('請在左側上傳檔案以開始。')