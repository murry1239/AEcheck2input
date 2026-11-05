"""ACI 결과 1행(H열 O) → ICH E2B(R3) ICSR XML.

요소 순서는 ICH 예시 2-1(자발 보고)과 Reference Instance v3.1을 따른다.
정보가 없는 항목은 지어내지 않고 nullFlavor로 두며, 담당자가 채울 항목은 issues로 돌려준다.
"""
import re
import uuid
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta

HL7 = 'urn:hl7-org:v3'
XSI = 'http://www.w3.org/2001/XMLSchema-instance'
ET.register_namespace('', HL7)
ET.register_namespace('xsi', XSI)

ICH = '2.16.840.1.113883.3.989.2.1'
OBS_CODE = ICH + '.1.19'
MEDDRA = '2.16.840.1.113883.6.163'
COUNTRY = '1.0.3166.1.2.2'
SEX = '1.0.5218'
# MFDS 지역(KR) 코드 체계 — cubeSAFETY 출력 XML에서 확인한 OID
KR_CAUSALITY_RESULT = '2.16.840.1.113883.3.989.5.1.10.1.5'
KR_CAUSALITY_METHOD = '2.16.840.1.113883.3.989.5.1.10.1.4'
CAUSALITY_EN = {'확실함': 'Certain', '상당히 확실함': 'Probable', '가능함': 'Possible', '평가 불능': 'Unassessable'}
SUMMARY_TODO = '이상사례 내용 확인 필요'  # ACI가 F열 유해사례 행에 넣는 문구
# E.i.3.2a~f: 사망, 생명 위협, 입원/연장, 장애, 선천 이상, 기타 의학적으로 중요
SERIOUSNESS = ('34', '21', '33', '35', '12', '26')
REPORTER = {'1': '의사', '2': '약사', '3': '보건의료전문가', '4': '법률가', '5': '소비자'}  # C.2.r.4
# 서술에서 뺄 VOC 머리말·상담 인사말
CHANNEL = re.compile(r'^\(.{1,20}\)$')  # 맨 위의 (홈페이지), (아웃콜) 같은 접수 경로
DROP_LINE = re.compile(r'^(?:최초\s*(?:인입|접수)일|답변\s*방법|이메일|전화\s*번호|고객\s*명|제목)(?:\s|[:：]|$)'
                       r'|^(?:문의|답변)\s*[:：)]'
                       r'|^(?:문의|답변)\)?$|^안녕하세요,?\s*일동제약입니다\.?$|^감사합니다\.?$'
                       r'|홈페이지를 방문해 주셔서|담당 부서로 전달하여|이메일로 답변드린 것으로')
ARROW = re.compile(r'^\s*(?:->|=>|→|>)\s*')


def sae_criterion(value):
    """J열 → E.i.3.2 순번(0=a 사망 … 5=f 기타). NI·빈칸 → None, 형식을 알 수 없으면 -1."""
    text = str(value or '').strip()
    if not text or text.upper() == 'NI':
        return None
    m = re.match(r'([a-fA-F])\s*[.)]', text)
    return 'abcdef'.index(m.group(1).lower()) if m else -1


def clean_story(text):
    """G열 원문에서 VOC 머리말·인사말을 빼고 상담 답변(-> …)은 ‘- ’ 목록으로 바꾼다."""
    out, in_answer = [], False
    for line in str(text or '').replace('\r\n', '\n').replace('\r', '\n').split('\n'):
        raw = line.rstrip()
        stripped = raw.strip()
        if not stripped:
            in_answer = False
            out.append('')
            continue
        if DROP_LINE.search(stripped) or (CHANNEL.match(stripped) and not any(out)):
            continue
        stripped = re.sub(r'^문의\s*내용\s*[:：]?\s*', '', stripped)
        if ARROW.match(raw):
            in_answer = True
            out.append('- ' + ARROW.sub('', raw).strip())
        elif in_answer and raw[:1] in (' ', '\t'):
            out.append('- ' + stripped)
        else:
            in_answer = stripped.startswith('- ')
            out.append(stripped)
    return re.sub(r'\n{3,}', '\n\n', '\n'.join(out)).strip()

# ACI 결과 열 순서 (A~M)
COLS = ('kind', 'date', 'sex', 'age', 'product', 'inquiry', 'text', 'ae', 'cube', 'sae', 'relatedness', 'qc', 'comment')


def is_ae(value):
    return str(value or '').strip()[:1] in ('O', 'o', '○', '◯')


def ae_terms(value):
    """'O — 복통, 설사' → ['복통', '설사']"""
    text = re.sub(r'^[Oo○◯]\s*[—–\-:：]?\s*', '', str(value or '').strip())
    terms = [t.strip() for t in re.split(r'[,;/·\n]', text)]
    return list(dict.fromkeys(t for t in terms if t))


def to_date(value, fallback=''):
    """'2026-10-26' 또는 Excel 날짜 숫자 → 'YYYYMMDD'"""
    text = str(value or '').strip()
    m = re.search(r'(\d{4})[-./]?(\d{2})[-./]?(\d{2})', text)
    if m:
        return ''.join(m.groups())
    if re.fullmatch(r'\d{5}(\.\d+)?', text):
        return (datetime(1899, 12, 30) + timedelta(days=int(float(text)))).strftime('%Y%m%d')
    return fallback


def to_sex(value):
    text = str(value or '')
    return '1' if '남' in text else '2' if '여' in text else ''


def age_band(value):
    """IDBI D열의 연령대. '70대', '70대이상', 그리고 숫자만 있는 '70'·'60'(10 단위)도 연령대로 본다 → 70."""
    text = str(value or '').replace(' ', '')
    m = re.match(r'(\d{1,3})대', text)
    if m and '미만' not in text:
        return int(m.group(1))
    m = re.fullmatch(r'(\d{1,3})', text)
    if m and int(m.group(1)) % 10 == 0 and int(m.group(1)) > 0:
        return int(m.group(1))
    return None


def to_age(value):
    """연령대 → (decade, '10.a'): IG D.2.2 ‘7 decades = 60대’이므로 70대 → 8. '55'·'55세' → (55, 'a')."""
    band = age_band(value)
    if band is not None:
        return band // 10 + 1, '10.a'
    m = re.fullmatch(r'(\d{1,3})(?:세)?', str(value or '').replace(' ', ''))
    if m:
        return int(m.group(1)), 'a'
    return None


def age_text(value):
    """서술용: '70' → '70대', '70대이상' → 그대로, '55' → '55세'."""
    text = str(value or '').strip()
    if re.fullmatch(r'\d{1,3}', text):
        return f'{text}대' if age_band(text) is not None else f'{text}세'
    return text


def case_id(cfg, report_date, row, cube=''):
    """I열(Cube) 사례번호가 있으면 cubeSAFETY 형식(KR-ID-261102SP_MAMTA_0)으로, 없으면 임시 번호."""
    cube = re.sub(r'\s+', '', str(cube or ''))
    if cube:
        if not cube.upper().startswith(cfg['country'] + '-'):
            cube = cfg['case_prefix'] + cube
        if not re.search(r'_\d+$', cube):
            cube += cfg['case_version_suffix']
        return cube
    return f"{cfg['case_prefix']}{report_date}SP-{row:03d}"  # 임시 번호: KR-ID-20261103SP-050


def age_group(age):
    """D.2.3: 범위가 한 그룹에 들어갈 때만. 20~50대 → 5 성인, 70대 이상 → 6 노인."""
    band = age_band(age)
    if band is None:
        return ''
    return '5' if 20 <= band <= 50 else '6' if band >= 70 else ''


def stable_uuid(*parts):
    return str(uuid.uuid5(uuid.NAMESPACE_URL, '/'.join(map(str, parts))))


def el(parent, tag, text=None, xsi_type=None, **attrs):
    e = ET.SubElement(parent, f'{{{HL7}}}{tag}')
    if xsi_type:
        e.set(f'{{{XSI}}}type', xsi_type)
    for key, value in attrs.items():
        e.set(key, str(value))
    if text is not None:
        e.text = text
    return e


def narrative(c, cfg):
    """H.1: 요약 머리말 + ‘70대 남성 소비자가 보고함.’으로 시작하는 정리된 원문."""
    sex = {'1': '남성', '2': '여성'}.get(to_sex(c['sex']), '')
    age = age_text(c['age'])  # '70' → '70대', '55' → '55세'
    lines = [f"접수일 {c['date_text']}",
             f"원보고자: 성별 {sex or '미기재'}, 나이 {age or '미기재'}",
             f"의심 제품: {c['product'] or '미기재'}",
             f"이상사례: {', '.join(c['terms']) or '미기재'}"]
    if sae_criterion(c['sae']) not in (None, -1):
        lines.append(f"중대성: {c['sae']}")
    if c['relatedness']:
        lines.append(f"회사 인과관계: {c['relatedness']}")
    who = ' '.join(x for x in (age, sex) if x)
    reporter = REPORTER.get(cfg.get('reporter_qualification', ''), '보고자')
    opening = f"{who} {reporter}가 보고함." if who else f"{reporter}가 보고함."
    story = clean_story(c['text'])
    return '\n'.join(lines) + '\n\n' + (opening + ' ' + story if story else opening)


def build(c, cfg, meddra, now=None):
    """c: parse_case() 결과. 반환: (XML bytes, issues)"""
    now = now or datetime.now()
    stamp = now.strftime('%Y%m%d%H%M%S') + cfg['timezone']
    cid = c['case_id']
    issues = [] if c['cube'] else ['I열(Cube) 사례번호가 없어 임시 번호를 썼습니다. cubeSAFETY 사례번호를 I열에 넣어 주세요.']
    criterion = sae_criterion(c['sae'])
    version = meddra.get('meddra_version', '').strip()
    codes = meddra.get('terms', {})

    root = ET.Element(f'{{{HL7}}}MCCI_IN200100UV01', {'ITSVersion': 'XML_1.0',
                      f'{{{XSI}}}schemaLocation': 'urn:hl7-org:v3 MCCI_IN200100UV01.xsd'})
    el(root, 'id', extension=f"{cfg['sender_id']}-{now:%Y%m%d%H%M%S}-{c['row']:04d}", root=ICH + '.3.22')  # N.1.2
    el(root, 'creationTime', value=stamp)  # N.1.5
    el(root, 'responseModeCode', code='D')
    el(root, 'interactionId', extension='MCCI_IN200100UV01', root='2.16.840.1.113883.1.6')
    el(root, 'name', code='1', codeSystem=ICH + '.1.1')  # N.1.1 ichicsr

    msg = el(root, 'PORR_IN049016UV')
    el(msg, 'id', extension=cid, root=ICH + '.3.1')  # N.2.r.1
    el(msg, 'creationTime', value=stamp)  # N.2.r.4
    el(msg, 'interactionId', extension='PORR_IN049016UV', root='2.16.840.1.113883.1.6')
    el(msg, 'processingCode', code='P')
    el(msg, 'processingModeCode', code='T')
    el(msg, 'acceptAckCode', code='AL')
    device = el(el(msg, 'receiver', typeCode='RCV'), 'device', classCode='DEV', determinerCode='INSTANCE')
    el(device, 'id', extension=cfg['receiver_id'], root=ICH + '.3.12')  # N.2.r.3
    device = el(el(msg, 'sender', typeCode='SND'), 'device', classCode='DEV', determinerCode='INSTANCE')
    el(device, 'id', extension=cfg['sender_id'], root=ICH + '.3.11')  # N.2.r.2

    act = el(msg, 'controlActProcess', classCode='CACT', moodCode='EVN')
    el(act, 'code', code='PORR_TE049016UV', codeSystem='2.16.840.1.113883.1.18')
    el(act, 'effectiveTime', value=stamp)  # C.1.2
    inv = el(el(act, 'subject', typeCode='SUBJ'), 'investigationEvent', classCode='INVSTG', moodCode='EVN')
    el(inv, 'id', extension=cid, root=ICH + '.3.1')  # C.1.1
    el(inv, 'id', extension=cid, root=ICH + '.3.2')  # C.1.8.1
    el(inv, 'code', code='PAT_ADV_EVNT', codeSystem='2.16.840.1.113883.5.4')
    story = narrative(c, cfg)
    el(inv, 'text', story)  # H.1
    el(inv, 'statusCode', code='active')
    el(el(inv, 'effectiveTime'), 'low', value=c['report_date'])  # C.1.4
    el(inv, 'availabilityTime', value=c['report_date'])  # C.1.5

    # 환자·이상사례·약물
    assessment = el(el(inv, 'component', typeCode='COMP'), 'adverseEventAssessment', classCode='INVSTG', moodCode='EVN')
    role = el(el(assessment, 'subject1', typeCode='SBJ'), 'primaryRole', classCode='INVSBJ')
    person = el(role, 'player1', classCode='PSN', determinerCode='INSTANCE')
    el(person, 'name', nullFlavor='UNK')  # D.1: VOC에서 환자 이니셜을 따로 뽑지 않음
    sex = to_sex(c['sex'])
    if sex:
        el(person, 'administrativeGenderCode', code=sex, codeSystem=SEX)  # D.5
    else:
        el(person, 'administrativeGenderCode', nullFlavor='UNK')
    age = to_age(c['age'])
    if age:
        obs = el(el(role, 'subjectOf2', typeCode='SBJ'), 'observation', classCode='OBS', moodCode='EVN')
        el(obs, 'code', code='3', codeSystem=OBS_CODE)
        el(obs, 'value', xsi_type='PQ', value=age[0], unit=age[1])  # D.2.2a/b
    group = age_group(c['age'])
    if group:
        obs = el(el(role, 'subjectOf2', typeCode='SBJ'), 'observation', classCode='OBS', moodCode='EVN')
        el(obs, 'code', code='4', codeSystem=OBS_CODE)
        el(obs, 'value', xsi_type='CE', code=group, codeSystem=ICH + '.1.9')  # D.2.3
    history = el(el(role, 'subjectOf2', typeCode='SBJ'), 'organizer', classCode='CATEGORY', moodCode='EVN')
    el(history, 'code', code='1', codeSystem=ICH + '.1.20')
    obs = el(el(history, 'component', typeCode='COMP'), 'observation', classCode='OBS', moodCode='EVN')
    el(obs, 'code', code='18', codeSystem=OBS_CODE)
    el(obs, 'value', xsi_type='ED', nullFlavor='UNK')  # D.7.2

    terms = c['terms'] or ['이상사례 내용 미기재']
    if not c['terms']:
        issues.append('H열에 이상사례 내용이 없습니다.')
    reaction_ids = []
    for i, term in enumerate(terms, 1):
        rid = stable_uuid(cid, 'reaction', i)
        reaction_ids.append(rid)
        obs = el(el(role, 'subjectOf2', typeCode='SBJ'), 'observation', classCode='OBS', moodCode='EVN')
        el(obs, 'id', root=rid)
        el(obs, 'code', code='29', codeSystem=OBS_CODE)
        code = str(codes.get(term, '')).strip()
        if code and version:
            value = el(obs, 'value', xsi_type='CE', code=code, codeSystem=MEDDRA, codeSystemVersion=version)  # E.i.2.1a/b
        else:
            value = el(obs, 'value', xsi_type='CE', nullFlavor='NI')
            issues.append(f"MedDRA 코드 필요: ‘{term}’" + ('' if code else ' (meddra_map.json에 없음)'))
        el(value, 'originalText', term, language=cfg['language'])  # E.i.1.1a/b
        place = el(el(el(obs, 'location', typeCode='LOC'), 'locatedEntity', classCode='LOCE'),
                   'locatedPlace', classCode='COUNTRY', determinerCode='INSTANCE')
        el(place, 'code', code=cfg['country'], codeSystem=COUNTRY)  # E.i.9
        for index, code_ in enumerate(SERIOUSNESS):  # E.i.3.2a~f: J열 기준만 true, 나머지 NI
            sub = el(el(obs, 'outboundRelationship2', typeCode='PERT'), 'observation', classCode='OBS', moodCode='EVN')
            el(sub, 'code', code=code_, codeSystem=OBS_CODE)
            if index == criterion:
                el(sub, 'value', xsi_type='BL', value='true')
            else:
                el(sub, 'value', xsi_type='BL', nullFlavor='NI')
        sub = el(el(obs, 'outboundRelationship2', typeCode='PERT'), 'observation', classCode='OBS', moodCode='EVN')
        el(sub, 'code', code='27', codeSystem=OBS_CODE)
        el(sub, 'value', xsi_type='CE', code='0', codeSystem=ICH + '.1.11')  # E.i.7: Unknown
    if c['terms'] and SUMMARY_TODO in c['terms']:
        issues.append('H열 이상사례 내용이 ‘확인 필요’ 상태입니다. 실제 이상사례로 고쳐 주세요.')
    if criterion == -1:
        issues.append(f"J열(SAE) 값 ‘{c['sae']}’을 알 수 없어 중대성을 넣지 않았습니다. NI 또는 a~f 중 하나로 고쳐 주세요.")
    elif criterion is not None:
        issues.append(f"중대한 이상사례({c['sae']})입니다. 신속보고 대상(C.1.7 = true)으로 넣었으니 보고 기한을 확인하세요.")

    drugs = el(el(role, 'subjectOf2', typeCode='SBJ'), 'organizer', classCode='CATEGORY', moodCode='EVN')
    el(drugs, 'code', code='4', codeSystem=ICH + '.1.20')
    drug_id = stable_uuid(cid, 'drug', 1)
    drug = el(el(drugs, 'component', typeCode='COMP'), 'substanceAdministration', classCode='SBADM', moodCode='EVN')
    el(drug, 'id', root=drug_id)
    product = el(el(el(drug, 'consumable', typeCode='CSM'), 'instanceOfKind', classCode='INST'),
                 'kindOfProduct', classCode='MMAT', determinerCode='KIND')
    el(product, 'name', c['product'] or '제품명 미기재')  # G.k.2.2
    if not c['product']:
        issues.append('E열 품목명이 없습니다.')

    role_ = el(el(assessment, 'component', typeCode='COMP'), 'causalityAssessment', classCode='OBS', moodCode='EVN')
    el(role_, 'code', code='20', codeSystem=OBS_CODE)
    el(role_, 'value', xsi_type='CE', code='1', codeSystem=ICH + '.1.13')  # G.k.1: Suspect
    el(el(el(role_, 'subject2', typeCode='SUBJ'), 'productUseReference', classCode='SBADM', moodCode='EVN'), 'id', root=drug_id)
    if c['relatedness'] and c['relatedness'] not in CAUSALITY_EN:
        issues.append(f"K열 값 ‘{c['relatedness']}’은 인과관계 4개 범주가 아니어서 넣지 않았습니다.")
    kr_codes = cfg.get('kr_causality_codes', {})
    for rid in reaction_ids:  # G.k.9.i.2.r: cubeSAFETY처럼 보고자(not reported) + 회사 평가
        assess(assessment, rid, drug_id, 'Reporter', text='not reported')
        if c['relatedness'] in CAUSALITY_EN:
            if cfg.get('regional') == 'KR' and kr_codes.get(c['relatedness']):
                assess(assessment, rid, drug_id, cfg['causality_source'], code=kr_codes[c['relatedness']],
                       method=cfg['kr_causality_method_code'])
            else:
                assess(assessment, rid, drug_id, cfg['causality_source'],
                       text=f"{CAUSALITY_EN[c['relatedness']]} ({c['relatedness']})", method_text=cfg['causality_method'])
    notes = [f'{label}: {c[key]}' for key, label in (('qc', 'QC'), ('comment', 'comment')) if c[key]]
    if notes:  # H.4
        comment = el(el(assessment, 'component1', typeCode='COMP'), 'observationEvent', classCode='OBS', moodCode='EVN')
        el(comment, 'code', code='10', codeSystem=OBS_CODE)
        el(comment, 'value', ' / '.join(notes), xsi_type='ED')
        el(el(el(comment, 'author', typeCode='AUT'), 'assignedEntity', classCode='ASSIGNED'), 'code', code='1', codeSystem=ICH + '.1.21')

    expedited = 'true' if criterion not in (None, -1) else 'false'  # 중대한 사례는 국내 신속보고(15일) 대상
    for code_, value in (('1', 'false'), ('23', expedited)):  # C.1.6.1, C.1.7
        obs = el(el(inv, 'component', typeCode='COMP'), 'observationEvent', classCode='OBS', moodCode='EVN')
        el(obs, 'code', code=code_, codeSystem=OBS_CODE)
        el(obs, 'value', xsi_type='BL', value=value)
    # H.1이 이미 한국어이므로 H.5(원어 요약)는 cubeSAFETY 출력처럼 넣지 않는다.

    first = el(inv, 'outboundRelationship', typeCode='SPRT')  # C.1.8.2
    rel = el(first, 'relatedInvestigation', classCode='INVSTG', moodCode='EVN')
    el(rel, 'code', code='1', codeSystem=ICH + '.1.22')
    author = el(el(el(rel, 'subjectOf2', typeCode='SUBJ'), 'controlActEvent', classCode='CACT', moodCode='EVN'), 'author', typeCode='AUT')
    el(el(author, 'assignedEntity', classCode='ASSIGNED'), 'code', code='2', codeSystem=ICH + '.1.3')

    source = el(inv, 'outboundRelationship', typeCode='SPRT')  # C.2.r
    el(source, 'priorityNumber', value='1')  # C.2.r.5
    rel = el(source, 'relatedInvestigation', classCode='INVSTG', moodCode='EVN')
    el(rel, 'code', code='2', codeSystem=ICH + '.1.22')
    author = el(el(el(rel, 'subjectOf2', typeCode='SUBJ'), 'controlActEvent', classCode='CACT', moodCode='EVN'), 'author', typeCode='AUT')
    reporter = el(el(author, 'assignedEntity', classCode='ASSIGNED'), 'assignedPerson', classCode='PSN', determinerCode='INSTANCE')
    qualification = el(reporter, 'asQualifiedEntity', classCode='QUAL')
    if cfg.get('reporter_qualification'):
        el(qualification, 'code', code=cfg['reporter_qualification'], codeSystem=ICH + '.1.6')  # C.2.r.4
    else:
        el(qualification, 'code', nullFlavor='UNK')
    place = el(el(reporter, 'asLocatedEntity', classCode='LOCE'), 'location', classCode='COUNTRY', determinerCode='INSTANCE')
    el(place, 'code', code=cfg['country'], codeSystem=COUNTRY)  # C.2.r.3

    sender = el(el(el(inv, 'subjectOf1', typeCode='SUBJ'), 'controlActEvent', classCode='CACT', moodCode='EVN'),
                'author', typeCode='AUT')
    entity = el(sender, 'assignedEntity', classCode='ASSIGNED')
    el(entity, 'code', code='1', codeSystem=ICH + '.1.7')  # C.3.1: 제약사
    if cfg.get('sender_state'):
        el(el(entity, 'addr'), 'state', cfg['sender_state'])  # C.3.4.3
    if cfg.get('sender_email'):
        el(entity, 'telecom', value='mailto:' + cfg['sender_email'])  # C.3.4.8
    if cfg.get('sender_given_name'):
        person_ = el(entity, 'assignedPerson', classCode='PSN', determinerCode='INSTANCE')
        el(el(person_, 'name'), 'given', cfg['sender_given_name'])  # C.3.3.3
        place = el(el(person_, 'asLocatedEntity', classCode='LOCE'), 'location', classCode='COUNTRY', determinerCode='INSTANCE')
        el(place, 'code', code=cfg['country'], codeSystem=COUNTRY)  # C.3.4.5
    dept = el(entity, 'representedOrganization', classCode='ORG', determinerCode='INSTANCE')
    if cfg.get('sender_department'):
        el(dept, 'name', cfg['sender_department'])  # C.3.3.1
    org = el(el(dept, 'assignedEntity', classCode='ASSIGNED'), 'representedOrganization', classCode='ORG', determinerCode='INSTANCE')
    el(org, 'name', cfg['sender_organisation'])  # C.3.2

    for code_, system, value in (('1', '.1.2', '1'), ('2', None, None)):  # C.1.3 자발 보고, C.1.9.1
        char = el(el(inv, 'subjectOf2', typeCode='SUBJ'), 'investigationCharacteristic', classCode='OBS', moodCode='EVN')
        el(char, 'code', code=code_, codeSystem=ICH + '.1.23')
        if system:
            el(char, 'value', xsi_type='CE', code=value, codeSystem=ICH + system)
        else:
            el(char, 'value', xsi_type='BL', nullFlavor='NI')

    device = el(el(root, 'receiver', typeCode='RCV'), 'device', classCode='DEV', determinerCode='INSTANCE')
    el(device, 'id', extension=cfg['receiver_id'], root=ICH + '.3.14')  # N.1.4
    device = el(el(root, 'sender', typeCode='SND'), 'device', classCode='DEV', determinerCode='INSTANCE')
    el(device, 'id', extension=cfg['sender_id'], root=ICH + '.3.13')  # N.1.3

    if not c['report_date']:
        issues.append('접수일(B열)을 알 수 없습니다.')
    ET.indent(root, space='  ')
    return b'<?xml version="1.0" encoding="UTF-8"?>\n' + ET.tostring(root, encoding='utf-8'), issues


def assess(assessment, reaction_id, drug_id, source, text='', code='', method='', method_text=''):
    """G.k.9.i.2.r 인과관계 평가 1건. code는 KR 지역 WHO-UMC 결과 코드."""
    cause = el(el(assessment, 'component', typeCode='COMP'), 'causalityAssessment', classCode='OBS', moodCode='EVN')
    el(cause, 'code', code='39', codeSystem=OBS_CODE)
    if code:
        el(cause, 'value', xsi_type='CE', code=code, codeSystem=KR_CAUSALITY_RESULT)  # G.k.9.i.2.r.3.KR.1
    else:
        el(cause, 'value', text, xsi_type='ST')  # G.k.9.i.2.r.3
    if method:
        el(cause, 'methodCode', code=method, codeSystem=KR_CAUSALITY_METHOD)  # G.k.9.i.2.r.2.KR.1
    elif method_text:
        el(el(cause, 'methodCode'), 'originalText', method_text)  # G.k.9.i.2.r.2
    el(el(el(el(cause, 'author', typeCode='AUT'), 'assignedEntity', classCode='ASSIGNED'), 'code'), 'originalText', source)
    el(el(el(cause, 'subject1', typeCode='SUBJ'), 'adverseEffectReference', classCode='OBS', moodCode='EVN'), 'id', root=reaction_id)
    el(el(el(cause, 'subject2', typeCode='SUBJ'), 'productUseReference', classCode='SBADM', moodCode='EVN'), 'id', root=drug_id)


def parse_case(values, row, file_date, cfg, anonymize):
    v = dict(zip(COLS, (str(x or '').strip() for x in list(values) + [''] * 13)))
    report_date = to_date(v['date'], file_date)
    text, _ = anonymize(v['text'])
    return dict(row=row, case_id=case_id(cfg, file_date or report_date, row, v['cube']), report_date=report_date,
                date_text=v['date'] or '미기재', kind=v['kind'], sex=v['sex'], age=v['age'], product=v['product'],
                inquiry=v['inquiry'], text=text, terms=ae_terms(v['ae']), sae=v['sae'],
                relatedness=v['relatedness'], cube=v['cube'], qc=v['qc'], comment=v['comment'])
