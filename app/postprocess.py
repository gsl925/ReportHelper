# app/postprocess.py
import re
import json
from rapidfuzz import process, fuzz


def load_domain_dict(path='domain_dict.json'):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def clean_text(text: str) -> str:
    # 基本清理：統一換行、去除多重空白
    t = text.replace('\r', '\n')
    t = re.sub('\n{2,}', '\n', t)
    t = re.sub('[ \t]{2,}', ' ', t)
    return t.strip()


def extract_key_sentences(text: str, domain_dict: dict, topn: int = 5):
    # 以關鍵字抽句（非常簡單的 heuristic）
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    scored = []
    for ln in lines:
        score = 0
        for kw in domain_dict.get('keywords', []):
            if kw in ln:
                score += 2
        for code in domain_dict.get('product_codes', []):
            if code in ln:
                score += 3
        scored.append((score, ln))
    scored.sort(reverse=True, key=lambda x: x[0])
    return [s for sc, s in scored[:topn]]


def simple_star_from_sentences(sentences: list):
    # 將 sentences 分配到 STAR（非常簡化的 rules）
    situation = []
    task = []
    action = []
    result = []
    for s in sentences:
        low = s.lower()
        if any(k in low for k in ['fail', 'error', '異常', 'fail rate', 'good rate', '良率']):
            situation.append(s)
        elif any(k in low for k in ['需要', '需', '目標', '目標是', '要求']):
            task.append(s)
        elif any(k in low for k in ['已', '已經', '調整', '更改', '修正', '採取']):
            action.append(s)
        elif any(k in low for k in ['回升', '改善', '改善為', '結果', '暫時']):
            result.append(s)
        else:
            # fallback: put into action if has verb
            action.append(s)
    return {
        'situation': situation,
        'task': task,
        'action': action,
        'result': result
    }


def apply_domain_corrections(text: str, domain_dict: dict):
    # 嘗試校正常見錯字或替換成 domain term
    for code in domain_dict.get('product_codes', []):
        # fuzzy replace
        matches = process.extract(code, [text], scorer=fuzz.partial_ratio, limit=1)
        # not implementing full replace here - placeholder
    return text