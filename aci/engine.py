"""Local candidate screening; evidence stays separate from missing information."""
import hashlib
import json
import re
from pathlib import Path

from anonymize import anonymize

ROOT = Path(__file__).resolve().parent
RULES = json.loads((ROOT / 'rules.json').read_text(encoding='utf-8'))
CATEGORIES = RULES['categories']
FACT_KEYS = ('temporal', 'collected', 'dechallenge', 'rechallenge')


def fingerprint(values):
    return hashlib.sha256(json.dumps(values[:7], ensure_ascii=False, separators=(',', ':')).encode()).hexdigest()


def map_causality(facts):
    """Only user's four combinations; ? never becomes X."""
    t, c, d, r = (facts.get(k, '?') for k in FACT_KEYS)
    if t == 'O' and d == 'O' and r == 'O':
        return '확실함'
    if t == 'O' and d == 'O' and r == 'X':
        return '상당히 확실함'
    if t == 'O' and c == 'O' and d == 'X' and r == 'X':
        return '가능함'
    if t == 'O' and c == 'X' and d == 'X' and r == 'X':
        return '평가 불능'
    return ''


def screen(values, row, sheet):
    text = str(values[6] or '')
    kind = str(values[5] or '')
    facts = dict.fromkeys(FACT_KEYS, '?')
    terms = [s for s in RULES['symptoms'] if s in text]
    headings = re.findall(r'(?:부작용|이상사례|유해사례|이상반응)\s*[:：]?\s*\(([^)]+)\)', text)
    exposure = bool(re.search(RULES['exposure_pattern'], text, re.S))
    hypo = bool(re.search(RULES['hypothetical_pattern'], text))
    neg = bool(re.search(RULES['explicit_negative_pattern'], text))
    actual_onset = bool(re.search(r'(?:발생했|발생하였|발생함|발생[.!]|생겼|나타났|증상이 있음|증상 있음)', text))
    recovered_history = bool(re.search(r'(?:중단|끊|호전|소실|없어짐|사라짐)', text))
    reasons = []
    if headings and (not hypo or actual_onset):
        ae = 'O'
        summary = '; '.join(dict.fromkeys(headings))
        reasons.append('불만 원문에 부작용/이상사례와 증상 기록이 있습니다.')
    elif terms and exposure and (not hypo or actual_onset) and (not neg or recovered_history):
        ae = 'O'
        summary = ', '.join(terms)
        reasons.append('사용·복용 문맥과 증상 문구를 함께 확인한 후보입니다.')
    else:
        ae = 'X'
        summary = ''
        if terms or '유해사례 등의 안전성' in kind or re.search('부작용|이상반응|이상사례', text):
            reasons.append('증상/안전성 단서가 있지만 실제 발생 여부를 원문으로 확인해야 합니다.')
        else:
            reasons.append('규칙에서 AE 단서를 찾지 못했습니다. 새로운 표현 누락 여부를 검토하세요.')
    if ae == 'O':
        if exposure and not hypo:
            facts['temporal'] = 'O'
        if re.search(RULES['collected_pattern'], text):
            facts['collected'] = 'O'
        dechallenge = re.search(RULES['dechallenge_pattern'], text, re.S)
        if dechallenge and not re.search(r'(?:않|못|아니|없어지지|호전되지)', dechallenge.group()):
            facts['dechallenge'] = 'O'
        rechallenge = re.search(RULES['rechallenge_pattern'], text, re.S)
        if rechallenge and not re.search(r'(?:않|못|아니|없었|없음)', rechallenge.group()):
            facts['rechallenge'] = 'O'
    category = map_causality(facts) if ae == 'O' else ''
    if ae == 'O' and not category:
        reasons.append('인과관계 기준을 확인할 정보가 부족하여 K열을 비워 둡니다.')
    return dict(sheet=sheet, row=row, key=fingerprint(values), ae=ae, summary=summary,
                facts=facts, relatedness=category, reason=' '.join(reasons),
                reviewed=False, method='규칙 초안')


def analyze(workbook):
    """AE 판단은 원문으로 하고, 화면·결과 파일에는 G열 익명화 문구만 넘긴다."""
    seed_path = ROOT / 'sample_review.json'
    seeds = json.loads(seed_path.read_text(encoding='utf-8')) if seed_path.exists() else {}
    records = []
    for sheet in workbook['sheets']:
        for row in sheet['records']:
            result = screen(row['values'], row['row'], sheet['name'])
            if result['key'] in seeds:
                seed = {k: v for k, v in seeds[result['key']].items() if k != 'snippets'}
                result.update(seed)
                result['method'] = '원문 검토 초안'
            values = list(row['values'])
            values[6], result['anonymized'] = anonymize(values[6])
            result['values'] = values
            records.append(result)
    return records


def validate_records(records, source_records):
    source = {(r['sheet'], r['row']): r for r in source_records}
    if len(records) != len(source):
        raise ValueError('원본과 검토 행 수가 다릅니다.')
    seen = set()
    for r in records:
        identity = (r['sheet'], r['row'])
        if identity not in source or identity in seen:
            raise ValueError('행이 중복되거나 원본에 없는 행입니다.')
        seen.add(identity)
        original = source[identity]
        if r.get('key') != original['key']:
            raise ValueError('원본 행 식별값이 다릅니다. 다시 불러오세요.')
        if r.get('ae') not in ('O', 'X'):
            raise ValueError('AE는 O 또는 X여야 합니다.')
        if r.get('relatedness', '') not in ['', *CATEGORIES]:
            raise ValueError('K열 값이 허용된 네 가지 범주에 없습니다.')
        if r['ae'] == 'X' and r.get('relatedness'):
            raise ValueError('AE가 X인 행의 K열은 비워 주세요.')
        if r['ae'] == 'O' and not r.get('summary', '').strip():
            raise ValueError('AE가 O인 행에 이상사례 내용을 입력하세요.')
        if len(r.get('summary', '')) > 32000:
            raise ValueError('AE 요약이 Excel 셀 길이 제한을 넘습니다.')
        if any(v not in ('O', 'X', '?') for v in r.get('facts', {}).values()):
            raise ValueError('인과관계 근거 값은 O, X, 미확인만 사용할 수 있습니다.')
        # G열은 화면에서 보낸 값이 아니라 서버가 만든 익명화 문구를 쓴다.
        r['g_text'] = str(original['values'][6] or '')
    return records
