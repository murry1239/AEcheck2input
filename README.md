# AEcheck2input — IAXG (IDBI AE XML Generator)

IDBI VOC 불만 자료에서 이상사례(AE)를 확인해 엑셀로 정리하고(**ACI**), AE 사례를 안전성 DB용 ICH E2B(R3) XML로 만드는(**A2X**) 업무 자동화 도구입니다. 두 단계를 **IAXG** 한 화면에서 씁니다.

| 단계 | 화면 탭 | 입력 | 결과 (작업 폴더 아래) | 버전 |
| --- | --- | --- | --- | --- |
| 1 | **1. ACI** · AE 검토 | `IDBI_YYYYMMDD.xls` | `ACI_outputs\IDBI_YYYYMMDD_ACI.xls`, `…_검토기록.json` | 0.7 |
| 2 | **2. A2X** · AE → XML | ACI 결과 엑셀 | `A2X_outputs\ICSR_R3(KR)_사례번호_제품명.xml`, `…_A2X결과.html/json` | 0.4 |

- ACI에서 결과 엑셀을 저장하면 A2X 탭에 그 파일이 골라지고, **2. A2X에서 Narration 검토 →** 버튼이 나타납니다 (원시자료 정리봇의 A → B 흐름과 같음).
- A2X는 이전 이름이 AAI(AE Auto Input)였습니다. 안전성 DB(cubeSAFETY, Argus, LSMV, Veeva Vault Safety 등)에 직접 접속하지 않고, E2B(R3) 가져오기로 읽을 XML 파일을 만듭니다.

## 실행

Windows, 데스크톱 MS Office Excel, Python 3.10 이상이 필요합니다.

1. `IAXG_실행.cmd`를 더블클릭합니다. 브라우저에 IAXG 화면이 열립니다.
   - **바탕화면 바로가기**: 처음 한 번 `IAXG_바로가기_만들기.cmd`를 더블클릭하면 바탕화면에 **IAXG** 아이콘이 생깁니다. 이후에는 그 아이콘으로 실행합니다. 프로그램 폴더를 옮기면 다시 만들어 주세요. 바로가기를 마우스 오른쪽 버튼으로 눌러 작업 표시줄·시작 메뉴에 고정할 수도 있습니다.
   - 브라우저 즐겨찾기는 쓸 수 없습니다. IAXG는 실행할 때마다 주소(포트)와 접근 토큰이 바뀌고, 즐겨찾기로는 프로그램을 시작할 수 없기 때문입니다. 실행하면 화면이 자동으로 열립니다.
2. **1. ACI** 탭: 작업 폴더를 지정하고 `IDBI_YYYYMMDD.xls`를 불러와 행별로 AE·SAE·인과관계를 검토합니다. AE가 O인 행은 I열 사례번호가 `KR-ID-YYMMDDSP_행번호`로 채워지며 고칠 수 있습니다. 끝나면 **결과 엑셀 저장**.
3. **2. A2X** 탭: 저장한 ACI 결과가 골라져 있습니다. **Narration 검토** → 사례별 서술(H.1)을 고치고 ‘검토 필요’ 사항을 확인 → **XML 만들기**로 XML·결과 보고서를 저장합니다. 검토 단계에서는 파일을 만들지 않습니다.
4. 화면의 주의사항대로 AE명(MedDRA)과 서술을 확인·수정하고, 안전성 DB에 가져온 뒤 확인·QC 합니다.
5. 끝나면 **도구 종료** (서버가 끝나고 탭이 닫힘). 오류는 **로그기록**에서 확인합니다 (`logs/iaxg_YYYYMMDD.log`).

자세한 내용: [ACI_사용안내.md](ACI_사용안내.md), [A2X_작업명세.md](A2X_작업명세.md), 지금까지의 요구사항: [프롬프트_정리.md](프롬프트_정리.md). A2X의 XSD 검증에는 ICH E2B(R3) IG 패키지 폴더 `E2B(R3)_IG_Complete _package v1_11`가 필요합니다 (저장소에는 포함하지 않음).

## 구성

- `IAXG_실행.cmd`, `Start-IAXG.ps1` — 실행 파일, `IAXG_바로가기_만들기.cmd`, `Make-IAXG-Shortcut.ps1`, `iaxg/iaxg.ico` — 바탕화면 바로가기
- `iaxg/app.py`, `iaxg/index.html` — IAXG 화면 서버 (127.0.0.1 전용). `/`는 1·2단계 탭, `/aci`·`/a2x`는 각 단계 화면
- `aci/app.py`, `aci/index.html` — ACI 화면·API
- `aci/engine.py`, `aci/rules.json` — AE 후보·SAE 항목·인과관계 규칙
- `aci/anonymize.py` — G열 이름·전화번호·주소·이메일 익명화 (A2X도 같이 씀)
- `aci/excel.ps1`, `a2x/read_aci.ps1` — Excel 읽기·저장·저장 후 검증
- `a2x/app.py`, `a2x/index.html` — A2X 화면·API (`/api/a2x/…`)
- `a2x/main.py`, `a2x/e2b.py` — ACI 결과 → E2B(R3) XML, XSD 검증, 결과 보고서
- `a2x/a2x_config.json` — 발송자/수신자 ID 등 설정, `a2x/meddra_map.json` — 한국어 표현 → MedDRA LLT 코드(담당자 입력)
- `tests/` — 단위 테스트 (`python -m unittest discover -s tests`)

불만 원문, 결과 엑셀·XML, 로그는 개인정보가 포함될 수 있어 저장소에 올리지 않습니다.
