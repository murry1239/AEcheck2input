import shutil
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT / 'aai'), str(ROOT / 'aci')]
from e2b import ae_terms, age_group, case_id, clean_story, is_ae, narrative, parse_case, sae_criterion, to_age, to_date, to_sex  # noqa: E402
from anonymize import anonymize  # noqa: E402
import main  # noqa: E402

CFG = main.CONFIG


def row(n, ae, text='복용 후 복통', k='', sex='여자', age='60대', cube='', sae='NI'):
    return {'row': n, 'values': ['불만', '2026-10-26', sex, age, '테스트정 10MG', '기타', text, ae, cube, sae, k, '', '']}


def observation_values(xml):
    """관찰 코드 → BL 값 목록 (value 또는 nullFlavor)."""
    ns = '{urn:hl7-org:v3}'
    found = {}
    for obs in ET.fromstring(xml).iter():
        code, value = obs.find(ns + 'code'), obs.find(ns + 'value')
        if code is not None and value is not None and value.get('{http://www.w3.org/2001/XMLSchema-instance}type') == 'BL':
            found.setdefault(code.get('code'), []).append(value.get('value') or value.get('nullFlavor'))
    return found


def workbook(*rows):
    return {'name': 'IDBI_20261026_ACI.xls', 'source_hash': 'x', 'sheets': [{'name': 's', 'records': list(rows)}]}


class ParseTests(unittest.TestCase):
    def test_age_uses_ig_decade_rule(self):
        self.assertEqual(to_age('40대'), (5, '10.a'))  # IG: 7 decades = 60대
        self.assertEqual(to_age('70대이상'), (8, '10.a'))
        self.assertEqual(to_age('55'), (55, 'a'))
        self.assertIsNone(to_age(''))
        self.assertIsNone(to_age('10대미만'))
        self.assertEqual([age_group(a) for a in ('20대', '50대', '60대', '70대이상', '10대', '55')], ['5', '5', '', '6', '', ''])

    def test_ae_column(self):
        self.assertTrue(is_ae('O — 복통, 설사'))
        self.assertFalse(is_ae('X'))
        self.assertFalse(is_ae(''))
        self.assertEqual(ae_terms('O — 복통, 설사; 복통'), ['복통', '설사'])
        self.assertEqual(ae_terms('O'), [])

    def test_date_and_sex(self):
        self.assertEqual(to_date('2026-10-26'), '20261026')
        self.assertEqual(to_date('46321'), '20261026')  # Excel 날짜 숫자
        self.assertEqual(to_date('', '20261001'), '20261001')
        self.assertEqual((to_sex('남성'), to_sex('여자'), to_sex('')), ('1', '2', ''))

    def test_case_id_from_cube_column(self):
        self.assertEqual(case_id(CFG, '20261102', 5, '261102SP_MAMTA'), 'KR-ID-261102SP_MAMTA_0')
        self.assertEqual(case_id(CFG, '20261102', 5, 'KR-ID-261023SP_FAC_0'), 'KR-ID-261023SP_FAC_0')
        self.assertEqual(case_id(CFG, '20261102', 5, 'KR-ID-261023SP_FAC'), 'KR-ID-261023SP_FAC_0')
        self.assertEqual(case_id(CFG, '20261103', 50, ''), 'KR-ID-20261103SP-050')  # 임시 번호

    def test_sae_criterion(self):
        self.assertEqual([sae_criterion(v) for v in ('NI', '', 'a. 사망', 'c. 입원 또는 입원기간의 연장', 'f. 기타', 'O')],
                         [None, None, 0, 2, 5, -1])

    def test_narrative_matches_requested_format(self):
        text = ('이전에 골드 복용했을 때에는 피로회복 느끼는게 확실히 느껴졌는데, 지금은 3개월 복용해도 별 느낌이 없는 것 같음. '
                '성분이 바뀌였는지?\n\n-> 골드 현재 단종되었지만, 성분 동일함\n    의약품임. 효능효과 개인차 있을 수 있음 ')
        c = parse_case(['불만', '2026-11-03', '남자', '70대', '아로나민골드정', '기타', text, 'O — 효과 없는 약물', '', 'NI',
                        '평가 불능', '', ''], 50, '20261103', CFG, anonymize)
        self.assertEqual(narrative(c, CFG),
                         '접수일 2026-11-03\n원보고자: 성별 남성, 나이 70대\n의심 제품: 아로나민골드정\n이상사례: 효과 없는 약물\n'
                         '회사 인과관계: 평가 불능\n\n70대 남성 소비자가 보고함. 이전에 골드 복용했을 때에는 피로회복 느끼는게 확실히 '
                         '느껴졌는데, 지금은 3개월 복용해도 별 느낌이 없는 것 같음. 성분이 바뀌였는지?\n\n'
                         '- 골드 현재 단종되었지만, 성분 동일함\n- 의약품임. 효능효과 개인차 있을 수 있음')

    def test_story_drops_voc_headers_and_greetings(self):
        home = ('(홈페이지)\n최초인입일 20261025\n\n문의)\n고객명 : HKD\n이메일 : a***@b.com\n답변방법 : 이메일\n제목 혈압계\n'
                '문의내용 혈압계 사용 후 두통이 있음\n\n답변)\n안녕하세요, 일동제약입니다.\n당사 홈페이지를 방문해 주셔서 감사합니다.\n감사합니다.')
        self.assertEqual(clean_story(home), '혈압계 사용 후 두통이 있음')
        self.assertEqual(clean_story('복통\n-> 확인 중\n    (답변 후 문자 예정)'), '복통\n- 확인 중\n- (답변 후 문자 예정)')


class ConvertTests(unittest.TestCase):
    def run_convert(self, wb, meddra):
        out = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, out, True)
        return main.convert(wb, out, meddra), out

    def xml(self, s, out, i=0):
        return (out / s['cases'][i]['file']).read_text(encoding='utf-8')

    def test_only_h_column_o_rows_become_xml(self):
        s, out = self.run_convert(workbook(row(2, 'O — 복통', cube='261026SP_TST'), row(3, 'X'), row(4, '')),
                                  {'meddra_version': '', 'terms': {}})
        self.assertEqual((s['ae_cases'], s['skipped_rows']), (1, 2))
        self.assertEqual([p.name for p in out.glob('*.xml')], ['ICSR_R3(KR)_KR-ID-261026SP_TST_0_테스트정10MG.xml'])

    def test_missing_meddra_or_cube_is_review_and_complete_is_done(self):
        s, _ = self.run_convert(workbook(row(2, 'O — 복통', cube='261026SP_TST')), {'meddra_version': '', 'terms': {}})
        self.assertEqual(s['cases'][0]['status'], '검토 필요')
        # 12345678은 테스트용 가짜 코드
        coded = {'meddra_version': '29.1', 'terms': {'복통': '12345678'}}
        s, out = self.run_convert(workbook(row(2, 'O — 복통')), coded)
        self.assertEqual(s['cases'][0]['issues'], ['I열(Cube) 사례번호가 없어 임시 번호를 썼습니다. cubeSAFETY 사례번호를 I열에 넣어 주세요.'])
        s, out = self.run_convert(workbook(row(2, 'O — 복통', cube='261026SP_TST')), coded)
        self.assertEqual(s['cases'][0]['status'], '완료')
        self.assertIn('code="12345678" codeSystem="2.16.840.1.113883.6.163" codeSystemVersion="29.1"', self.xml(s, out))

    def test_xml_content(self):
        text = '홍길동(010-1234-5678) 복용 후 복통'
        s, out = self.run_convert(workbook(row(2, 'O — 복통', text, k='가능함', cube='261026SP_TST')), {'meddra_version': '', 'terms': {}})
        xml = self.xml(s, out)
        self.assertNotIn('홍길동', xml)
        self.assertNotIn('1234-5678', xml)
        self.assertIn('<id extension="KR-ID-261026SP_TST_0" root="2.16.840.1.113883.3.989.2.1.3.1" />', xml)
        self.assertIn('<value xsi:type="ST">not reported</value>', xml)
        self.assertIn('<value xsi:type="CE" code="3" codeSystem="2.16.840.1.113883.3.989.5.1.10.1.5" />', xml)
        self.assertIn('<value xsi:type="PQ" value="7" unit="10.a" />', xml)
        self.assertIn('<originalText language="kor">복통</originalText>', xml)
        self.assertIn('<administrativeGenderCode code="2" codeSystem="1.0.5218" />', xml)
        self.assertIn('<code code="5" codeSystem="2.16.840.1.113883.3.989.2.1.1.6" />', xml)
        self.assertIn('extension="MFDS-2020000084"', xml)

    def test_sae_sets_criterion_and_expedited(self):
        s, out = self.run_convert(workbook(row(2, 'O — 복통, 설사', cube='T', sae='c. 입원 또는 입원기간의 연장')),
                                  {'meddra_version': '', 'terms': {}})
        flags = observation_values(self.xml(s, out))
        self.assertEqual(flags['33'], ['true', 'true'])  # E.i.3.2c, 이상사례 2개 모두
        self.assertEqual(flags['34'], ['NI', 'NI'])  # E.i.3.2a
        self.assertEqual(flags['23'], ['true'])  # C.1.7 신속보고
        self.assertIn('중대성: c. 입원 또는 입원기간의 연장', s['cases'][0]['narrative'])
        s, out = self.run_convert(workbook(row(2, 'O — 복통', cube='T')), {'meddra_version': '', 'terms': {}})
        flags = observation_values(self.xml(s, out))
        self.assertEqual((flags['33'], flags['23']), (['NI'], ['false']))

    def test_outputs_go_to_one_folder(self):
        s, out = self.run_convert(workbook(row(2, 'O — 복통', cube='T')), {'meddra_version': '', 'terms': {}})
        self.assertEqual(sorted(p.name for p in out.iterdir()),
                         ['ICSR_R3(KR)_KR-ID-T_0_테스트정10MG.xml', 'IDBI_20261026_AAI결과.html', 'IDBI_20261026_AAI결과.json'])

    def test_xsd_valid(self):
        if not (ROOT / CFG['schema_path']).exists():
            self.skipTest('ICH E2B(R3) 패키지가 없음')
        rows = [row(2, 'O — 복통, 설사', k='확실함', cube='A'), row(3, 'O', sex='', age='', cube='B'),
                row(5, 'O — 두통', k='모름', age='70대이상', cube='C')]
        s, _ = self.run_convert(workbook(*rows), {'meddra_version': '29.1', 'terms': {'복통': '12345678'}})
        self.assertEqual(s['schema'], '검증함')
        self.assertEqual([c['xsd_error'] for c in s['cases']], ['', '', ''])
        self.assertIn('H열에 이상사례 내용이 없습니다.', s['cases'][1]['issues'])
        self.assertTrue(any('4개 범주가 아니어서' in i for i in s['cases'][2]['issues']))


if __name__ == '__main__':
    unittest.main()
