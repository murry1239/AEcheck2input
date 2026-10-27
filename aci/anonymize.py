"""G열 개인정보 익명화: 이름→영문 이니셜, 전화번호→앞 한 자리, 주소→도로명/동까지, 이메일→앞 한 글자."""
import re

# 성(姓) 영문 표기의 첫 글자. 표에 없는 성은 아래 자모 규칙을 따른다.
SURNAME = {
    '김': 'K', '이': 'L', '박': 'P', '최': 'C', '정': 'J', '강': 'K', '조': 'C', '윤': 'Y', '장': 'J', '임': 'L',
    '한': 'H', '오': 'O', '서': 'S', '신': 'S', '권': 'K', '황': 'H', '안': 'A', '송': 'S', '류': 'R', '유': 'Y',
    '전': 'J', '홍': 'H', '고': 'K', '문': 'M', '양': 'Y', '손': 'S', '배': 'B', '백': 'B', '허': 'H', '남': 'N',
    '심': 'S', '노': 'N', '하': 'H', '곽': 'K', '성': 'S', '차': 'C', '주': 'J', '우': 'W', '구': 'K', '민': 'M',
    '진': 'J', '나': 'N', '지': 'J', '엄': 'U', '채': 'C', '원': 'W', '천': 'C', '방': 'B', '공': 'K', '현': 'H',
    '함': 'H', '변': 'B', '염': 'Y', '여': 'Y', '추': 'C', '도': 'D', '소': 'S', '석': 'S', '선': 'S', '설': 'S',
    '마': 'M', '길': 'K', '연': 'Y', '위': 'W', '표': 'P', '명': 'M', '기': 'K', '반': 'B', '왕': 'W', '금': 'K',
    '옥': 'O', '육': 'Y', '인': 'I', '맹': 'M', '제': 'J', '모': 'M', '탁': 'T', '국': 'K', '어': 'E', '은': 'E',
    '편': 'P', '용': 'Y', '예': 'Y', '봉': 'B', '경': 'K', '사': 'S', '부': 'B', '태': 'T', '목': 'M', '형': 'H',
}
# 초성 19자 순서. ㄱ은 예시(홍길동→HKD)에 맞춰 K.
INITIAL = 'KKNDDRMBBSSOJJCKTPH'
# 초성이 ㅇ이면 중성(21자 순서)으로 첫 글자를 정한다.
VOWEL = 'AAYYEEYYOWWOYUWWWYEUI'
STOPWORDS = {
    '연락처', '전화번호', '휴대폰', '핸드폰', '연락번호', '대표번호', '담당자', '고객님', '고객', '번호', '전화',
    '문의', '안내', '확인', '유선', '팩스', '본인', '보호자', '환자', '부모', '어머니', '아버지', '기타', '없음',
    '미상', '모름', '약국', '병원', '의원', '회사', '본사', '지점', '대리점', '공급처', '구매처', '구입처',
}
SURNAMES = ''.join(SURNAME)
NAME = '[' + SURNAMES + '][가-힣]{1,2}'
LABEL = (r'(?:보내는\s?분|받는\s?분|고객\s?명|고객\s?성함|고객|성명|성함|이름|환자\s?명|환자|보호자|신고자|'
         r'문의자|작성자|담당자|수령인|주문자|요청자)')
PHONE = re.compile(r'(?<!\d)(0\d{1,2})([-.)\s]?)(\d{3,4})([-.\s]?)(\d{4})(?!\d)')
EMAIL = re.compile(r'([A-Za-z0-9._%+-])[A-Za-z0-9._%+-]*@([A-Za-z0-9.-]+\.[A-Za-z]{2,})')
# 이름 뒤에 올 수 있는 것: 줄 끝, 괄호, 구분자, 호칭, 전화번호. "담당자 부재로"처럼 문장이 이어지면 이름으로 보지 않는다.
NAME_END = r'(?=[ \t]*(?:$|[(\[/,]|님|씨|0\d))'
DEPT = r'[가-힣A-Za-z0-9&]+(?:부|팀|파트|센터|부서)'
NAME_BY_LABEL = re.compile(LABEL + r'(\s*[:：]\s*(?:' + DEPT + r'\s+)?)(' + NAME + r')' + NAME_END, re.M)
NAME_BY_DEPT = re.compile(r'(?<![가-힣])(' + DEPT + r'\s+)(' + NAME + r')' + NAME_END, re.M)
NAME_BY_PHONE = re.compile(r'(?<![가-힣])(' + NAME + r')(?=\s*[(\[]?\s*0\d{1,2}[-.)\s]?\d{3,4})')
NAME_BY_TITLE = re.compile(r'(?<![가-힣])(' + NAME + r')(?=\s?(?:사원|대리|과장|차장|부장|팀장|주임|매니저))')
REGION = (r'(?:서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충청|충북|충남|전라|전북|전남|경상|경북|경남|제주)'
          r'[가-힣]*')
DISTRICT = r'[가-힣]+(?:시|군|구|읍|면)'
AREA = r'(?:' + REGION + r'(?:\s+' + DISTRICT + r')*|' + DISTRICT + r'(?:\s+' + DISTRICT + r')*)'
# 도로명 주소: 도로명까지 남기고 번길·건물번호·상세주소(동·호·층)는 지운다.
ROAD_ADDRESS = re.compile(
    r'(' + AREA + r'\s+[가-힣]+(?:대로|로|길))'
    r'(?:\s*\d+(?:번)?길)?\s*\d+(?:-\d+)?(?:번지)?'
    r'(?:[^\n]{0,40}\d+\s*(?:호|층)(?![가-힣]))?')
# 지번 주소: 동·리·가까지 남기고 번지는 지운다.
LOT_ADDRESS = re.compile(r'(' + AREA + r'\s+[가-힣0-9]+(?:동|리|가))\s*\d+(?:-\d+)?(?:번지)?(?:[^\n]{0,40}\d+\s*(?:호|층)(?![가-힣]))?')


def initials(name):
    letters = [SURNAME.get(name[0]) or syllable_letter(name[0])]
    letters += [syllable_letter(ch) for ch in name[1:]]
    return ''.join(letters)


def syllable_letter(ch):
    code = ord(ch) - 0xAC00
    if not 0 <= code < 11172:
        return ch
    initial = code // 588
    if initial == 11:  # ㅇ
        return VOWEL[(code % 588) // 28]
    return INITIAL[initial]


def mask_name(match, group):
    name = match.group(group)
    if name in STOPWORDS:
        return match.group(0)
    start, end = match.span(group)
    whole = match.group(0)
    offset = match.start(0)
    return whole[:start - offset] + initials(name) + whole[end - offset:]


def mask_phone(match):
    head, sep1, middle, sep2, last = match.groups()
    sep1 = '-' if not sep1 or sep1 == ')' else sep1
    sep2 = sep2 or '-'
    return head + sep1 + middle[0] + 'xxxx' + sep2 + last[0] + 'xxxx'


def anonymize(text):
    """반환: (익명화 문자열, 처리 건수). 원문은 바꾸지 않는다."""
    text = str(text or '')
    count = 0

    def sub(pattern, repl, value):
        nonlocal count
        result, n = pattern.subn(repl, value)
        changed = result != value
        count += n if changed else 0
        return result

    # 이름은 전화번호를 지우기 전에 찾는다(이름 바로 뒤 전화번호가 단서).
    text = sub(NAME_BY_LABEL, lambda m: mask_name(m, 2), text)
    text = sub(NAME_BY_DEPT, lambda m: mask_name(m, 2), text)
    text = sub(NAME_BY_PHONE, lambda m: mask_name(m, 1), text)
    text = sub(NAME_BY_TITLE, lambda m: mask_name(m, 1), text)
    text = sub(EMAIL, lambda m: m.group(1) + '***@' + m.group(2), text)
    text = sub(PHONE, mask_phone, text)
    text = sub(ROAD_ADDRESS, lambda m: m.group(1), text)
    text = sub(LOT_ADDRESS, lambda m: m.group(1), text)
    return text, count
