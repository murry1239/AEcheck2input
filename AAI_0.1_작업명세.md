# AAI 0.1 작업명세 — ACI 결과 → ICH E2B(R3) XML

## 목표

ACI 결과 엑셀(`IDBI_YYYYMMDD_ACI.xls`)에서 **H열이 `O`인 행만** 사례 1건으로 보고, 사례마다 **ICH E2B(R3) ICSR XML** 파일을 하나씩 만든다. cubeSAFETY, Argus Safety, LSMV, Veeva Vault Safety 등의 E2B(R3) 가져오기(import)로 읽을 수 있는 형식을 목표로 한다.

- **웹을 쓰지 않는다.** ACI처럼 PC에서 엑셀을 읽고, 정해진 폴더에 XML 파일을 저장한다. 안전성 DB로 가져오는 작업은 사용자가 각 시스템의 E2B 가져오기 기능으로 한다.
- 저장 위치: ACI 작업 폴더(`aci_settings.json`의 workspace) 아래 `aai_outputs\IDBI_YYYYMMDD_ACI_시각\`. ACI 결과가 `outputs\`에 저장되는 것과 같은 방식.
- 파일 이름: cubeSAFETY 출력과 같은 꼴인 `ICSR_R3(KR)_사례번호_제품명.xml`. 같은 폴더에 결과 보고서 `AAI_결과.html`, `AAI_결과.json`.
- XML은 ICH XSD(`MCCI_IN200100UV01.xsd`)로 구조를 검증한다.
- XSD가 확인하지 않는 업무 규칙(MedDRA 코드, 사례번호 등)은 AAI가 따로 점검해 **검토 필요**로 보여 준다. 빠진 정보는 지어내지 않는다.

## 기준 자료

| 자료 | 용도 |
| --- | --- |
| `1_ICH_ICSR_Implementation_Guide_v5_03.pdf` | 데이터 요소 정의와 규칙의 기준 |
| `11_ICH E2B(R3) Core Data Elements and Business Rules ver.1.00.xlsx` | 요소별 필수 여부, nullFlavor 허용값 |
| `6_Example Instances/2-1_ExampleCase_newNum_initial_v1_0.xml` | 자발 보고 1건의 XML 구조 |
| `5_Reference Instances/00_ICH_ICSR_Reference_Instance_variation_v3_1.xml` | 2-1에 없는 요소 위치 (원어 이상사례명, 인과관계 등) |
| `7_E2B Bilingual Code Lists v2.11` | 코드값 |
| `국내외시판후의약품이상사례보고시스템사용자매뉴얼(의약품안전나라)_v3.1` (저장소에 올리지 않음) | KR 지역 항목(…KR.1) 코드값과 작성 규칙 |
| `xml파일예시_261102/` (2건, 저장소에 올리지 않음) | **cubeSAFETY가 출력한 KR 지역 E2B(R3)**. 발송/수신 ID, 사례번호 형식, KR 인과관계 코드, 보고자 자격, 연령군 입력 방식을 맞춤 |

cubeSAFETY 예시와 AAI 출력의 요소 경로를 비교하면, **AAI가 쓰는 요소는 모두 cubeSAFETY 출력에도 있다.** cubeSAFETY에만 있는 것은 VOC 엑셀에 없는 정보(투여 기간·용법·제형, 적응증, 조치, MFDS 제품코드, 이상사례 시작일, 의료인 확인)다.

## 처리 흐름

1. `AAI_실행.cmd` → ACI 결과 엑셀 선택 (ACI 작업 폴더의 `outputs`에서 시작).
2. Excel로 A~M열을 읽기 전용으로 읽는다.
3. H열이 `O`(또는 `○`)로 시작하는 행만 사례로 만든다. 나머지는 건너뛰고 수만 센다.
4. G열은 ACI 익명화 규칙을 한 번 더 적용한다 (ACI 0.2 이전 결과 대비).
5. 사례마다 XML 생성 → XSD 검증 → 업무 규칙 점검 → 저장.
6. 보고서를 브라우저로 연다. 실행 기록은 `logs/aai_YYYYMMDD.log`.

## 매핑

### 설정값 (`aai/aai_config.json`, cubeSAFETY 출력 기준)

| E2B 요소 | 값 | 비고 |
| --- | --- | --- |
| N.1.3 / N.2.r.2 발송자 ID | `MFDS-2020000084` | cubeSAFETY 출력과 같음 |
| N.1.4 / N.2.r.3 수신자 ID | `MFDS-O-KR` | 〃 |
| C.1.3 보고 유형 | 1 (자발 보고) | |
| C.1.6.1 / C.1.7 | false / false | J열(SAE)이 `O`이면 **검토 필요** |
| C.1.8.2 최초 발송자 | 2 (Other) | |
| C.1.9.1 | nullFlavor NI | |
| C.2.r.3 / E.i.9 국가 | KR | |
| C.2.r.4 보고자 자격 | 5 (소비자) | cubeSAFETY 예시와 같음. 설정에서 바꿀 수 있음 |
| C.3.1 / C.3.2 / C.3.3.1 / C.3.4.3 | 1 (제약사) / 일동제약 / PV team / Seoul | |
| C.3.3.3 / C.3.4.8 담당자 이름·이메일 | `aai/aai_config.local.json` | 개인 정보라 저장소에 올리지 않음 |

### ACI 결과 열 → E2B 요소

| ACI 열 | E2B 요소 | 변환 |
| --- | --- | --- |
| **I Cube** | **C.1.1 / C.1.8.1 / N.2.r.1 사례번호** | `261102SP_MAMTA` → `KR-ID-261102SP_MAMTA_0`. 이미 `KR-`로 시작하면 그대로, `_숫자`가 없으면 `_0`(최초 보고)을 붙임. **비어 있으면 임시 번호 `KR-ID-IDBI날짜-행번호` + 검토 필요** |
| B 일자 | C.1.4 최초 접수일, C.1.5 최신 정보일 | `2026-10-26` → `20261026` |
| C 성별 | D.5 | 남자/남성 → 1, 여자/여성 → 2, 그 외 nullFlavor UNK |
| D 나이 | D.2.2a/b, D.2.3 | `55` → 55 `a`. `40대` → 5 `10.a` (IG: “7 decades = 60대”). 연령군은 범위가 한 그룹일 때만: 20~50대 → 5(성인), 70대 이상 → 6(노인) |
| (없음) | D.1 환자 | nullFlavor UNK |
| (없음) | D.7.2 병력 | nullFlavor UNK |
| E 품목명 | G.k.2.2, G.k.1 = 1 (의심약) | |
| F 문의종류 | H.1 서술에 포함 | |
| G 불만사항 (익명화) | H.1 사례 서술 (한국어) | 접수 정보·환자·제품·이상사례·원문을 묶어 작성. H.1이 한국어이므로 H.5는 cubeSAFETY처럼 쓰지 않음 |
| H AE 내용 | E.i.1.1a 원어 이상사례명 (`kor`) | `O — 복통, 설사` → 이상사례 2개 |
| H AE 내용 | E.i.2.1a/b MedDRA | `aai/meddra_map.json`에서 찾음. 없으면 nullFlavor + **검토 필요** |
| J SAE | E.i.3.2a~f | 모두 nullFlavor NI. J열 `O`이면 **검토 필요** |
| K relatedness | G.k.9.i.2.r 인과관계 | cubeSAFETY처럼 이상사례마다 ① 보고자 `not reported` ② 회사 평가: KR 코드(방법 코드 1, 결과 코드 `kr_causality_codes`) |
| (없음) | E.i.7 결과 | 0 (Unknown) |
| L QC, M comment | H.4 발송자 의견 | 값이 있으면 옮김 |

### K열 → KR 인과관계 결과 코드

의약품안전나라 사용자 매뉴얼 v3.1 p.29(G.k.9.i.2.r.3.KR.1)로 확인했다. 평가방법(G.k.9.i.2.r.2.KR.1)은 국내보고이므로 1(WHO-UMC).

| K열 | 코드 | 매뉴얼 표기 |
| --- | --- | --- |
| 확실함 | 1 | 확실함 (Certain) |
| 상당히 확실함 | 2 | 상당히 확실함 (Probable/likely) |
| 가능함 | 3 | 가능함 (Possible) |
| 평가 불능 | 6 | 평가불가 (Unassessable/unclassifiable) |

ACI K열 기준에는 없는 4(가능성 적음), 5(평가곤란)는 쓰지 않는다.

### 의약품안전나라 매뉴얼에서 확인한 KR 규칙

- 최초 보고에서 C.1.8.1은 C.1.1과 같다 (p.16). AAI는 둘 다 I열 사례번호를 쓴다.
- 국내 보고의 제품코드(G.k.2.1.KR.1b)는 식약처 품목기준코드 (p.19). 0.1에서는 넣지 않는다.
- E.i.3.2a~f는 필수이며 정보가 없으면 nullFlavor NI (p.26).
- 테스트 보고는 N.1.4, N.2.r.3에 테스트 환경 수신자 정보를 넣는다 (별첨3 Q18). 의약품안전나라 > 통합자료실의 ‘개별 항목 검증 룰’에서 값을 확인해 `receiver_id`를 바꾸면 된다.

### MedDRA

- MedDRA는 라이선스 용어집이라 AAI가 코드를 만들어 내지 않는다.
- `aai/meddra_map.json`에 담당자가 한국어 표현과 LLT 코드, MedDRA 버전을 채운다. cubeSAFETY 예시는 29.0, 29.1을 썼다.
- 코드가 없는 이상사례는 원어 표현(E.i.1.1a)만 넣고 코드는 nullFlavor `NI`로 둔다. XSD는 통과하지만 **MedDRA 코드 필요**로 표시한다.

## 결과 판정

| 상태 | 뜻 |
| --- | --- |
| 완료 | XSD 통과, 점검 항목 없음 |
| 검토 필요 | XSD는 통과했지만 MedDRA 코드, I열 사례번호, SAE 확인 등 채울 항목이 있음 |
| 오류 | XSD 검증 실패. 가져오기에 쓰지 말 것 |

## 0.1 범위 밖 (다음 버전 후보)

- 안전성 DB에서 가져오기를 자동 실행하는 기능
- 후속 보고(follow-up, `_1`…), 무효화
- MedDRA 자동 코딩, 투여 기간·용법·적응증·조치 추출, MFDS 제품코드(G.k.2.1.KR.1)
- 여러 사례를 한 배치 파일로 묶기

## 남은 확인 사항

1. MedDRA 버전과 자주 쓰는 표현의 LLT 코드 (`aai/meddra_map.json`).
2. I열에 넣는 사례번호 형식 (`261102SP_MAMTA`처럼 앞부분만인지, `KR-ID-…_0` 전체인지). 두 형식 모두 처리한다.
