import sys
import unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'aci'))
from engine import map_causality, screen, fingerprint, validate_records, analyze
from anonymize import anonymize, initials

class EngineTests(unittest.TestCase):
    def test_user_table_and_missing_data(self):
        for collected in ['O','X','?']:
            f=dict(temporal='O',collected=collected,dechallenge='O',rechallenge='O')
            self.assertEqual(map_causality(f),'확실함')
            f['rechallenge']='X'
            self.assertEqual(map_causality(f),'상당히 확실함')
        self.assertEqual(map_causality(dict(temporal='O',collected='O',dechallenge='X',rechallenge='X')),'가능함')
        self.assertEqual(map_causality(dict(temporal='O',collected='X',dechallenge='X',rechallenge='X')),'평가 불능')
        self.assertEqual(map_causality(dict(temporal='O',collected='?',dechallenge='?',rechallenge='?')),'')
        self.assertEqual(map_causality(dict(temporal='X',collected='X',dechallenge='X',rechallenge='X')),'')

    def make(self,text,kind='기타'):
        return ['', '2026-10-01', '', '', '제품', kind, text]

    def test_other_inquiry_type_is_not_excluded(self):
        r=screen(self.make('20261001 / 제품 부작용(복통,설사)'),2,'s')
        self.assertEqual(r['ae'],'O')
        self.assertEqual(r['relatedness'],'')

    def test_hypothetical_and_preexisting(self):
        r=screen(self.make('피로 때문에 복용하려고 함. 부작용으로 설사가 있을까요?'),2,'s')
        self.assertEqual(r['ae'],'X')
        r=screen(self.make('감기 증상이 있어 추가 복용 간격 문의'),2,'s')
        self.assertEqual(r['ae'],'X')
        r=screen(self.make('복용하고 부작용 증상은 없음'),2,'s')
        self.assertEqual(r['ae'],'X')

    def test_safety_label_does_not_prove_event(self):
        r=screen(self.make('구매 가능한 약국 문의','유해사례 등의 안전성'),2,'s')
        self.assertEqual(r['ae'],'X')

    def test_resolved_event_remains_an_ae(self):
        r=screen(self.make('20261001 / 제품 부작용(발진). 복용 중단 후 현재 증상 없음.'),2,'s')
        self.assertEqual(r['ae'],'O')
        r=screen(self.make('복용 후 발진 발생. 부작용인지 궁금합니다.'),2,'s')
        self.assertEqual(r['ae'],'O')

    def test_no_false_rechallenge_or_dechallenge(self):
        r=screen(self.make('복용 후 복통. 복용 중단 후 증상이 없어지지 않음. 다시 복용 시작 예정.'),2,'s')
        self.assertNotEqual(r['facts']['dechallenge'],'O')
        self.assertNotEqual(r['facts']['rechallenge'],'O')
        self.assertEqual(r['relatedness'],'')
        r=screen(self.make('복용 후 복통. 복용 중단 후 증상 없어짐. 다시 복용하니 증상 재발.'),2,'s')
        self.assertEqual(r['facts']['dechallenge'],'O')
        self.assertEqual(r['facts']['rechallenge'],'O')
        self.assertEqual(r['relatedness'],'확실함')

    def test_row_integrity_and_x_gating(self):
        r=screen(self.make('제품 부작용(복통)'),2,'s');r['values']=self.make('제품 부작용(복통)')
        self.assertEqual(validate_records([r],[r]),[r])
        bad=dict(r,ae='X',relatedness='확실함')
        with self.assertRaises(ValueError):validate_records([bad],[r])
        with self.assertRaises(ValueError):validate_records([dict(r,key='wrong')],[r])
        with self.assertRaises(ValueError):validate_records([r,r],[r])
        with self.assertRaises(ValueError):validate_records([dict(r,summary='')],[r])

    def test_g_text_comes_from_server_not_client(self):
        r=screen(self.make('제품 부작용(복통)'),2,'s');r['values']=self.make('서버 문구')
        client=dict(r,values=self.make('바꾼 문구'))
        self.assertEqual(validate_records([client],[r])[0]['g_text'],'서버 문구')

    def test_anonymize_examples(self):
        self.assertEqual(initials('홍길동'),'HKD')
        self.assertEqual(anonymize('010-1234-4567')[0],'010-1xxxx-4xxxx')
        self.assertEqual(anonymize('01012344567')[0],'010-1xxxx-4xxxx')
        self.assertEqual(anonymize('주소 : 서울 서초구 바우뫼로27길 2')[0],'주소 : 서울 서초구 바우뫼로')
        self.assertEqual(anonymize('주소 : 경기도 가상시 동안구 예시대로12번길3 꿈아파트 101동 1201호')[0],'주소 : 경기도 가상시 동안구 예시대로')
        self.assertEqual(anonymize('- 고객명 : 홍길동 ')[0],'- 고객명 : HKD ')
        self.assertEqual(anonymize('담당자 : 서울1영업부 김철수')[0],'담당자 : 서울1영업부 KCS')
        self.assertEqual(anonymize('홍길동(010-1234-4567)')[0],'HKD(010-1xxxx-4xxxx)')
        self.assertEqual(anonymize('메일 abc123@example.com')[0],'메일 a***@example.com')

    def test_anonymize_leaves_ordinary_words(self):
        for text in ['담당자 부재로 대신 전달','고객명 전달','담당자 연락','- 고객 : 이후 진행은 영업담당자와 소통','담당자 연락처 안내']:
            self.assertEqual(anonymize(text)[0],text)

    def test_analyze_masks_g_but_screens_original(self):
        text='홍길동(010-1234-4567) 복용 후 발진 발생'
        rows=analyze({'sheets':[{'name':'s','records':[{'row':2,'values':self.make(text)}]}]})
        self.assertEqual(rows[0]['values'][6],'HKD(010-1xxxx-4xxxx) 복용 후 발진 발생')
        self.assertEqual(rows[0]['ae'],'O')
        self.assertEqual(rows[0]['key'],fingerprint(self.make(text)))
        self.assertNotIn('snippets',rows[0])

    def test_fingerprint_changes_with_source(self):
        self.assertNotEqual(fingerprint(self.make('복통')),fingerprint(self.make('설사')))

if __name__=='__main__':unittest.main()
