# AEcheck2input

IDBI VOC 불만 자료에서 이상사례(AE)를 확인해 엑셀로 정리하고(ACI), SafetyDB에 입력하는(AAI) 업무 자동화 도구입니다.

| 모듈 | 내용 | 상태 |
| --- | --- | --- |
| ACI · AE Check in IDBI | `IDBI_YYYYMMDD.xls`를 읽어 H~M열 추가, AE O/X·인과관계 초안, G열 개인정보 익명화 | 0.3 |
| AAI · AE Auto Input | 확인된 AE를 SafetyDB에 입력 | 설계 전 |

## ACI 실행

Windows, 데스크톱 MS Office Excel, Python 3.10 이상이 필요합니다.

1. `ACI_실행.cmd`를 더블클릭합니다. 브라우저에 검토 화면이 열립니다.
2. 작업 폴더를 지정하고 `IDBI_YYYYMMDD.xls`를 불러옵니다.
3. 행별로 AE·인과관계를 검토하고 **결과 엑셀 저장**을 누릅니다. 결과는 `outputs/` 아래에 저장되며 원본은 바뀌지 않습니다.

자세한 내용은 [ACI_사용안내.md](ACI_사용안내.md)를 참고하세요.

## 구성

- `aci/app.py` — 로컬 검토 화면 서버 (127.0.0.1 전용)
- `aci/engine.py`, `aci/rules.json` — AE 후보·인과관계 규칙
- `aci/anonymize.py` — G열 이름·전화번호·주소·이메일 익명화
- `aci/excel.ps1` — Excel 읽기·저장·저장 후 검증
- `tests/` — 단위 테스트 (`python -m unittest discover -s tests`)

불만 원문, 결과 엑셀, 로그는 개인정보가 포함될 수 있어 저장소에 올리지 않습니다.
