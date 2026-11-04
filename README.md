# AEcheck2input

IDBI VOC 불만 자료에서 이상사례(AE)를 확인해 엑셀로 정리하고(ACI), SafetyDB에 입력하는(AAI) 업무 자동화 도구입니다.

| 모듈 | 내용 | 상태 |
| --- | --- | --- |
| ACI · AE Check in IDBI | `IDBI_YYYYMMDD.xls`를 읽어 H~M열 추가, AE O/X·SAE·인과관계 초안, G열 개인정보 익명화 | 0.4 |
| AAI · AE Auto Input | ACI 결과의 AE(H열 O) 사례를 ICH E2B(R3) XML로 만들어 안전성 DB(cubeSAFETY, Argus, LSMV, Veeva Vault Safety 등)로 가져올 수 있게 함 | 0.2 |

두 도구는 같은 **작업 폴더**를 씁니다. ACI 결과는 `작업 폴더\ACI_outputs`, AAI 결과는 `작업 폴더\AAI_outputs`에 저장됩니다.

## ACI 실행

Windows, 데스크톱 MS Office Excel, Python 3.10 이상이 필요합니다.

1. `ACI_실행.cmd`를 더블클릭합니다. 브라우저에 검토 화면이 열립니다.
2. 작업 폴더를 지정하고 `IDBI_YYYYMMDD.xls`를 불러옵니다.
3. 행별로 AE·SAE·인과관계를 검토하고 **결과 엑셀 저장**을 누릅니다. 결과는 `ACI_outputs`에 `IDBI_YYYYMMDD_ACI.xls`, `IDBI_YYYYMMDD_검토기록.json`으로 저장되며 원본은 바뀌지 않습니다.

자세한 내용은 [ACI_사용안내.md](ACI_사용안내.md)를 참고하세요.

## AAI 실행

1. `AAI_실행.cmd`를 더블클릭합니다. ACI처럼 브라우저에 AAI 화면이 열립니다.
2. 작업 폴더를 확인하고, `ACI_outputs`의 ACI 결과(`IDBI_YYYYMMDD_ACI.xls`)를 골라 **XML 만들기**를 누릅니다.
3. H열이 `O`인 행만 사례 1건 = XML 1개로 만들어 `AAI_outputs`에 저장합니다. 사례번호는 I열(Cube) 값을 쓰고, 비어 있으면 임시 번호(`KR-ID-YYYYMMDDSP-행번호`)를 씁니다.
4. 화면에서 사례별 서술(H.1)과 ‘검토 필요’ 사항을 확인한 뒤 안전성 DB의 E2B(R3) 가져오기 기능으로 XML을 읽습니다. 안전성 DB에 직접 접속하지는 않습니다.

매핑과 판정 기준은 [AAI_0.2_작업명세.md](AAI_0.2_작업명세.md), 지금까지의 요구사항은 [프롬프트_정리.md](프롬프트_정리.md)를 참고하세요. XSD 검증에는 ICH E2B(R3) IG 패키지 폴더 `E2B(R3)_IG_Complete _package v1_11`가 필요합니다(저장소에는 포함하지 않음).

## 구성

- `aci/app.py` — 로컬 검토 화면 서버 (127.0.0.1 전용)
- `aci/engine.py`, `aci/rules.json` — AE 후보·인과관계 규칙
- `aci/anonymize.py` — G열 이름·전화번호·주소·이메일 익명화
- `aci/excel.ps1` — Excel 읽기·저장·저장 후 검증
- `aai/app.py`, `aai/index.html` — AAI 화면 서버 (127.0.0.1 전용)
- `aai/main.py`, `aai/e2b.py` — ACI 결과 → E2B(R3) XML, XSD 검증, 결과 보고서
- `aai/aai_config.json` — 발송자/수신자 ID 등 설정, `aai/meddra_map.json` — 한국어 표현 → MedDRA LLT 코드(담당자 입력)
- `tests/` — 단위 테스트 (`python -m unittest discover -s tests`)

불만 원문, 결과 엑셀, 로그는 개인정보가 포함될 수 있어 저장소에 올리지 않습니다.
