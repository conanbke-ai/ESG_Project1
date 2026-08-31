# 전국 재생에너지 표준화·태양광 발전량 예측·이상징후 알림 시스템

현재 확보한 한국남동발전·한국남부발전·한국동서발전·한국서부발전·한국농어촌공사 자료를 출발점으로, 기관을
고정하지 않고 공통 데이터 계약과 품질 게이트를 통과한 국내 발전소를 계속 추가하는 전국 발전량
예측 프로젝트입니다. 예측 잔차와 공개 기상자료는 이상징후 분석·알림에 사용합니다.

원본 표준화 계층은 태양광뿐 아니라 공개 파일에 함께 들어 있는 풍력·수력을 보존합니다. 태양광
모델은 `energy_source=solar`만 명시적으로 선택하며, 다른 발전원을 태양광 출력에 합치지 않습니다.

예측 모델에서의 `전국`은 공개 발전량 표본의 지역·기관 범위를 지속적으로 넓히는 목표이지 국내
모든 민간·공공 설비의 시간별 발전량을 이미 전수 확보했다는 뜻이 아닙니다. 전국 설비 현황 화면은
이 학습 모집단과 분리된 EPSIS 등록 범위를 사용합니다. 공개된 정비·고장 이력이 없으므로 설비
고장을 판정하거나 원인으로 확정하지 않습니다.

## 환경

- Python 3.11
- 주요 패키지: `pandas`, `numpy`, `torch`, `scikit-learn`, `xgboost`, `optuna`, `requests`, `selenium`

```powershell
.\.venv\Scripts\python.exe -m pip install -e ".[dev]"
```

## 전체 파이프라인 한 번에 실행

명시한 파일을 사용하는 경우:

```bash
python app.py pipeline "합산발전량(MWh)" --data file/merge_data/val.csv \
  --epochs 10 --n-trials 5 --optimizer-timeout-seconds 1800
```

파일을 생략하면 `--input-dir` 아래의 최신 CSV 또는 Excel 파일을 자동으로 선택합니다.

```bash
python app.py pipeline target_column --input-dir file/merge_data --features feature1,feature2
```

실행 순서는 다음과 같습니다.

1. 입력 파일 자동 탐색 및 로딩
2. 숫자형 피처 선택, 품질 flag/결측 마스크 생성, 전처리 CSV 저장
3. CNN-BiLSTM 모델 학습 및 체크포인트 저장
4. 평가와 이상징후 탐지
5. 최종 HTML 보고서와 실행 manifest 생성

각 실행 결과는 기본적으로 `output/pipeline/<실행시간>/`에 저장됩니다. 단계 사이의 DataFrame은 메모리로 전달하고 기본 `minimal` 모드에서는 최종 산출물만 저장합니다.

## 공식 데이터 자동 수집

현재 자동화된 4개 발전사 파일은 각 공식 홈페이지의 공공데이터 게시물 또는 그 게시물이 연결한
공공데이터포털 첨부파일에서 내려받습니다. 한국농어촌공사 영암 원본은 확보된 주기성 파일을
`prepare-data`가 자동 심사·편입합니다. 발전사별 다운로드 URL 환경변수는 필요하지 않습니다.
기상청 로그인이 필요한 경우에는 `KMA_CHROME_USER_DATA_DIR`, 중부발전 선택 수집에는
`DATA_GO_SERVICE_KEY`를 `.env.local`에 설정합니다. 형식은 [`.env.example`](.env.example)을 따릅니다.

```bash
python app.py collect --start-date 2024-01-01 \
  --sources koen,kospo,ewp,iwest,kma
```

새 발전사 CSV는 다운로드 시점 기준으로
`발전소명_[세부발전소명] 태양광발전실적_YYYYMMDD.csv`를 사용합니다. 예를 들어
`한국남부발전(주)_[남제주소내] 태양광발전실적_20250228.csv`입니다. 월별 통합 파일처럼 단일
발전소가 아닌 원본은 세부명에 `월간통합_YYYYMM`을 기록합니다. 재현 실행에서는
`--download-date 2025-02-28`로 명명 날짜를 고정할 수 있습니다.

보관한 모든 지원 발전사 원본을 공통 스키마로 바꾸고 학습 파일까지 만들려면 다음을 실행합니다.

```bash
python app.py prepare-data
```

`prepare-data`는 현재 `file/solar_data_file/`의 88개 원본을 4,236,565개 시간 행의 파일별 gzip CSV
파티션으로 변환합니다. 이어 농어촌공사 영암 6개 파일을 심사해 개체 식별이 가능한 2022~2025
4개 파일·105,191개 시간 행을 같은 계약으로 편입하고, 2020~2021은 식별자 부재로 격리합니다.
기본 원본 검사는 `generation_manifest.json`, 영암 검사는
`candidates/krc_yeongam/candidate_manifest.json`에 행 수·기간·결측·음수·중복·물리 상한과 함께 기록합니다.
각 파티션은 한 원본 파일 단위로 처리하고 완성된 임시 파일만 원자적으로 교체합니다. manifest에는
입력·출력 byte와 SHA-256을 기록해 같은 파일명의 수정본과 재처리 lineage를 추적합니다.
이후 공식 표준 발전량을 발전소·시간 단위로 재집계하고, 기존 병합본의 ASOS 지점번호는 승인
근거가 아닌 audit-only 후보로만 보존합니다. 공식 주소·좌표, 발전기간 전체를 덮는 KMA 지점 이력,
근거가 있는 reviewed mapping을 통과한 행으로 단일 호환본 `file/standardized/model_ready.csv.gz`와
학습용 회사×연도 파티션 `file/standardized/model_ready_parts/`를 생성합니다. 현재 Gold는
718,531행·24개 발전소(태양광 22, 풍력 2)이고, 태양광 43개와 소수력 1개를 포함한 44개 자산은
근거 없는 기상 결합 대신 registry에 격리합니다. 2026년 기상 연간 파일이 아직 없어 해당 발전량 2,832행·2개소는
Silver에는 보존하고 Gold에서만 보류합니다.
태양광 Gold 701,011행·22개소 가운데 여수태양광 29,280행은 일 총량이 고정 야간 버킷에 적재된
자료라 원문과 품질 근거는 보존하되 시간 예측 학습에서는 제외합니다. 따라서 현재 실제 태양광
학습 적격 범위는 671,731행·21개소입니다.
기존 병합 타깃과 공식 원본의
차이는 `legacy_pipeline_quality_report.csv`에 별도로 기록합니다. 생성물은
재현 가능한 중간 산출물이므로 Git에는 넣지 않습니다.

같은 실행에서 `plant_quality_report.csv`와 `quality_manifest.json`도 생성합니다. 발전소별 시간
커버리지, 음수, 용량초과, 주간 0, 양의 값 flatline, 기상 결측, 일별 형태 일관성, 지역 peer
상관을 기록하되 공개 데이터만으로 센서·설비 고장을 확정하지 않습니다.
일별 총량이 최소 30일 동안 95% 이상 한 개의 동일한 야간 시간 버킷에만 기록된 태양광 자료는
시간별 정답을 복원할 근거가 없으므로 `quality_daily_aggregate_profile`로 발전소 전체를 표시하고
시간 예측 학습에서 제외합니다. 일사가 짧은 날이나 일시 정지를 오인하지 않도록 네 조건을 모두
통과할 때만 적용하며, 원자료는 일 단위 집계 근거로 그대로 보존합니다.
현재 24개 Gold 발전소 중 `low` 20개, 사람 검토가 필요한 `review` 3개, `high` 1개입니다.
태양광은 `low` 20개·`review` 1개·`high` 1개이고 풍력 2개는 모두 `review`입니다. 전남은
태양광 8개소·273,955행과 풍력 2개소·17,520행, 전북은 태양광 2개소·22,824행입니다. 상세 원인과
제공자별 건수는 [`docs/DATA_COLLECTION_AUDIT.md`](docs/DATA_COLLECTION_AUDIT.md)에 있습니다.

`--end-date`를 생략하면 수집 종료일은 실행 당일입니다. 기상 관측자료는 당일 값이 불완전할 수
있어 전일까지, 남동발전은 현재 연도의 현재 월까지만 요청합니다.

- `koen`: 남동발전 발전량 조회 화면에서 요청 기간의 월별 CSV를 내려받습니다.
- `kospo`: 남부발전 공공데이터 목록이 연결한 데이터셋 `15156688`의 최신 첨부 ID를 조회해 CSV를 내려받습니다.
- `ewp`: 동서발전 공공데이터 게시물 `43582`의 전국 태양광 학습 CSV를 공식 POST 폼으로 내려받습니다.
- `iwest`: 서부발전 태양광 발전 현황 데이터셋 `15025486`의 최신 첨부 ID를 조회해 CSV를 내려받습니다.
- `komipo`: 중부발전 API 데이터셋 `15084511`을 본부코드×일자 단위 gzip Bronze 파티션으로
  제한 수집합니다. 공식 `daypower` 단위가 명시되지 않아 바로 MWh나 학습행으로 승격하지 않습니다.
- `kma`: 기상자료개방포털 ASOS 화면에서 시간자료를 최대 1년 단위로 나누어 CSV로 내려받습니다. API 키는 사용하지 않습니다.
- 요청 기간이 `file/KMA_data_file/OBS_ASOS_TIM_<연도>.csv`에 이미 포함되면 API를 호출하거나 파일을 복제하지 않고 기존 자료를 그대로 등록합니다. 현재 Gold 결합 가능 연도는 2013~2025입니다.
- 공공기관 기상관측 분 자료는 기관별 요소 차이와 1회 최대 1일 제한 때문에 기본 수집에서 제외했습니다.
- 기상청 로그인이 필요한 경우 `KMA_CHROME_USER_DATA_DIR`에 로그인 전용 Chrome 프로필을 지정합니다. 실행 중인 일반 Chrome 프로필과 같은 경로를 동시에 사용하지 않는 것을 권장합니다.
- 발전사 원본은 `file/raw/<회사>/`에 공급기관 byte 그대로 보존하고, 표준본은
  `file/standardized/downloads/<회사>/`에 Excel 호환 `UTF-8-SIG`로 원자 저장합니다.
- 공급기관 원본은 CP949와 UTF-8이 섞여 있는 것이 정상입니다. 원본을 일괄 재인코딩하면 hash와
  출처 재현성이 깨지므로 변경하지 않습니다. 새 collection manifest에는 파일별 인코딩·BOM·byte·
  SHA-256과 Bronze/Silver 역할을 기록합니다. 기존 `file/raw/**/normalized/`는 이전 실행의 파생본입니다.
- 과거 보관 원본은 당시 파일명과 hash를 유지합니다. 새 수집분부터 표준 명명 규칙을 강제하며,
  기존 파일을 보기 좋게 만들기 위한 일괄 rename은 lineage를 훼손하므로 수행하지 않습니다.
- 네 발전사 개별 설비 시간 발전량의 표준 계약은
  `timestamp, company, plant_id, plant, unit, energy_source, generation_mwh, capacity_mw, tilt_deg, latitude, longitude, address, source_file`입니다.
- `plant_id`에는 회사 코드를 접두사로 넣어 회사 간 같은 발전소명이 충돌하지 않게 했습니다.
- 원본에 설비용량·설치각·좌표·주소가 없으면 각 발전사의 공식 설비/사업현황 파일로 보강하고,
  근거가 불명확한 매칭은 결측으로 유지합니다.
- 원본 단위가 `Wh` 또는 `kWh`이면 정규화 단계에서 `MWh`로 변환합니다. `1시`는 해당 날짜의
  `00:00`, `24시`는 `23:00`으로 해석합니다.
- 남동발전의 일부 기존 파일은 kWh 스케일 값을 `MWh`로 잘못 표기했지만 최신 공식 파일에서는
  `KWh`로 정정됐습니다. 정규화기는 헤더와 시간 발전량 범위를 함께 검사해 이 레거시 오표기를 보정합니다.
- 남부발전 다송리 원본은 시간 버킷 헤더가 `kWh`지만 893kW 설비의 물리 상한과 일계 대조 결과
  `Wh`로 확인됐습니다. 선언 단위의 용량비 99백분위가 100을 넘고 1,000배 대안 하나만 최대
  용량비 1.2 이하인 경우에만 자동 교정하며, 선언·해석 단위와 일계 대조 결과를
  `generation_manifest.json`에 남깁니다. 일반적인 용량초과만으로 단위를 추정하지 않습니다.
- 동서발전 학습 표준본은 `timestamp + region`을 유일 키로 사용하며 기상·설비·시간 피처와
  `generation_mwh`를 포함합니다. 공식 파일에서 겹치는 2024년 수정 레코드는 뒤의 고정밀 레코드를 사용합니다.
- 서부발전 신재생에너지 파일의 풍력 17,520시간 행도 `energy_source=wind`로 보존합니다. 같은 방식으로
  명칭에서 명확히 식별되는 소수력 106,608시간 행은 `hydro`로 분류합니다. 태양광 학습 시에만
  발전원 필터를 적용합니다.
- 동일 일자·발전기 파일 행이 완전히 같으면 공식 파일의 단순 중복으로 제거하지만, 값이 다른 중복 키는
  임의로 선택하지 않고 오류로 중단합니다.
- 기상청 기존 파일 재사용 여부는 요청 기간과 `--overwrite`로 제어합니다. 같은 다운로드 날짜의
  발전사 원본이 이미 있으면 재사용하고, 공급기관 최신 첨부를 강제로 다시 받으려면 `--overwrite`를 사용합니다.
- 출처 하나가 실패해도 나머지 수집은 계속되며, 실패 출처에는 `error.log`가 생성됩니다.

중부발전 API 키를 발급받은 뒤에는 본부코드를 명시하고 호출 예산 안에서 탐색 수집합니다.

```powershell
python app.py collect --sources komipo --start-date 2026-08-01 --end-date 2026-08-07 `
  --komipo-station-codes CODE1,CODE2 --api-max-calls 100
```

완료한 station/day 파티션은 재호출하지 않으며 `--overwrite`일 때만 갱신합니다. API 결과는 단위,
설비 위치, 용량, 이력 연속성을 검토하기 전까지 Bronze 격리 상태입니다.

현재 확보 파일의 행 수·기간·결측·중복 검증 결과는
[`docs/DATA_COLLECTION_AUDIT.md`](docs/DATA_COLLECTION_AUDIT.md)에 기록했습니다.

## 로컬 대시보드

기존 루트별 정적 HTML과 모델 체크포인트 폴더를 분리했습니다. 웹 산출물의 단일 위치는
`dashboard/`이며, 세 화면은 같은 사용자용 JSON과 CSS/JS를 공유하되 역할을 명확히 분리합니다.

- `solar_dashboard.html`: 학습 데이터 보유 여부와 무관한 전국 태양광 발전설비 현황
- `forecast.html`: 발전소별 Test 구간의 실제 발전량과 선택 모델 예측 시계열
- `model_analysis.html`: 모델 비교, 지역·발전소별 성능, 예측 잔차와 데이터 이상 신호

과거 `plant_region_report_perm.html` 주소는 새 분석 화면으로 이동하는 호환 진입점만 유지합니다.

```powershell
python app.py build-dashboard
python app.py serve-dashboard --port 5500
```

공식 시도 경계를 다시 만들 때만 GIS 선택 의존성을 설치합니다. 다운로드 archive와 시도
Shapefile의 hash를 명시하고, 실제 변환에 사용한 `.shp/.shx/.dbf/.prj`가 검증된 archive 안의
동명 구성요소와 byte 단위로 같은지 확인합니다. archive·구성요소 hash가 다르거나 `BASE_DATE`가
기준일과 다르거나 17개 시도가 완전하지 않거나 대표 도서 소속이 틀리면 게시 파일을 교체하지
않습니다.

```powershell
.\.venv\Scripts\python.exe -m pip install -e ".[geo]"
python app.py prepare-boundaries `
  --source-shp "C:\path\to\bnd_sido_00_2025_2Q.shp" `
  --source-archive "C:\path\to\sgis_boundary_2025_2Q.zip" `
  --reference-date 2025-06-30 `
  --archive-sha256 f1cf0f9de453ac7eaacb273f39cee52851183372b9ddfda428a967c3a670b2c6 `
  --shapefile-sha256 8edb33f1f633002a41f9bc407943cbebf0c2d8d5b3c18bd3fa5eb41aabe7e77e
```

그다음 `http://127.0.0.1:5500/solar_dashboard.html`을 엽니다. 메인 화면은 한국전력거래소
EPSIS에서 `발전원=태양에너지`로 직접 다시 내려받은 2026-08-05 기준 전국 원본을 읽습니다.
합계행을 제외한 188,594개 공식 등록 레코드와 33,059.52 MW를 최신 16개 광역 기준으로 집계합니다.
해시·필수 컬럼뿐 아니라 모든 데이터 행의 `발전원=태양에너지` 계약을 검사하므로, 필터가
풀린 혼합 원본은 다른 발전원이 합산되기 전에 즉시 실패합니다.
원본의 320개 세부지역 표기는 시행일이 기록된 행정구역 reference와 검토된 행 단위 교정을 적용해
258개 표준 집계로 보존합니다. [행정안전부 행정표준코드관리시스템](https://www.code.go.kr/stdcode/regCodeL.do)의
2026-07-01 공식 변경을 기준으로 과거 `광주`·`전남`과 새 명칭을 `전남광주통합특별시`로
합치되 총 건수·용량은 바꾸지 않습니다. 이 명칭 통합은 충돌 행의 실제 소속을 이름만으로
추정하는 규칙이 아닙니다. 메인 지도는 세부지역 점을 한꺼번에 찍지 않고 광역
단계구분도만 보여줍니다. 경계는 국가데이터처 SGIS 2025년 2분기 공식 Shapefile을 WGS84로
변환한 결과이며, 원본 hash와 백령·대청·연평·울릉도 소속 검사를 통과해야 게시됩니다. 검색은
선택 시도에 갇히지 않고 시도명·약칭·세부지역명·검토된 자연지명을 전국에서 찾으며, 결과는
고정 높이 내부 스크롤 영역에서 갱신되어 입력 중 화면이 움직이지 않습니다. `백령도` 검색은
인천광역시 옹진군 전체 집계임을 명시하고, `울릉도`처럼 공식 행정구역은 있지만 EPSIS 세부지역
행이 없는 경우에는 설비를 임의 생성하지 않고 그 사실을 표시합니다. 세부지역 목록은 기본적으로
설비용량 높은 순이며, 등록건수·설비용량·세부지역명 각각에 대해 정방향과 역방향을 선택할 수
있습니다. 정렬 선택 상자뿐 아니라 표의 세 컬럼 헤더를 반복해서 눌러도 방향이 전환됩니다.
EPSIS 원천의 광역지역과 세부지역이 다른 범위는 51행·53.289725 MW입니다. 공식 근거를 검토해
확정한 47행·48.895470 MW에만 raw-row override를 적용합니다. 미해결 4행·4.394255 MW는 원본
광역지역 합계를 유지하되 표준 세부지역 집계에서 `미확정 지역`으로 격리합니다. Override와 격리
규칙은 `config/national_solar_location_overrides_20260828.json`에서 해당 원본의 SHA-256 및 그룹별
기대 행 수·설비용량 합계에 결합됩니다. 하나라도 달라지면 일부 결과를 게시하지 않고 전체 생성을
즉시 실패시킵니다. Bronze 원문과 전국 합계 188,594개·33,059.516180 MW는 이 처리로 바뀌지 않습니다.
행정구역·울릉도·독도·공식 허가주소 보강 판단은
[`docs/NATIONAL_REGION_REFERENCE_AUDIT.md`](docs/NATIONAL_REGION_REFERENCE_AUDIT.md)에 있습니다.
수동 발전소 별칭·좌표 캐시·기상관측소 매핑을 공식 자료로 교체할 수 있는 범위는
[`docs/OFFICIAL_SOURCE_REPLACEMENT_AUDIT.md`](docs/OFFICIAL_SOURCE_REPLACEMENT_AUDIT.md)에
정리했습니다.
이 행은 물리적 발전소 고유 식별자가 아닌 발전기·등록 레코드이므로 화면에서
`발전소 개소`라고 과장하지 않습니다. CP949 Bronze 원본, 다운로드 시각, 기준일, SHA-256,
제외 범위는 `config/national_solar_inventory.json`에서 버전 관리합니다.

`발전량 예측` 화면은 지역·발전소·모델을 선택해 정식 Test 구간의 실제값과 저장된 예측값을
같은 축에서 표시합니다. 이는 미래 운영예보가 아니며, 예보 발행시각이 보존된 기상예보 입력을
연결하기 전에는 미래 발전량을 생성하거나 화면에 꾸며내지 않습니다.

`모델 성능 분석` 화면은 XGBoost와 CNN-BiLSTM의 `timestamp + plant_id` 공통 Test 행을 다시
정렬한 경우에만 모델 우열과 지역·발전소 순위를 표시합니다. 모든 정식 실행을 평가 계약별로 묶어
가장 최신의 호환 XGBoost·CNN-BiLSTM 쌍을 고르며, 스모크 실행, 실패 실행, 서로 다른 평가기간,
구형·혼합 발전원 결과는 제외합니다. NMAE는 설비용량이 확인된 공통 표본으로 계산하고 적용률을
함께 표시합니다. 정식 공통 결과가 없으면 가상 수치나 0점 차트를 만들지 않고 빈 상태를 표시합니다.

예측 이상 신호의 기준은 Test와 분리된 Calibration에서 고정합니다. 발전소별 168개 이상 표본이
있으면 용량 정규화 잔차 기준을 사용하고, 표본이 부족하면 전역 용량 정규화 기준으로 후퇴합니다.
용량이 없는 발전소는 충분한 자체 Calibration 표본이 있을 때만 절대오차 기준을 사용합니다. 전체
신호 건수·지역·발전소 집계와 화면에 보관하는 모델별 상위 250개 대표 이벤트를 분리하며, 시계열은
발전소별 최근 연속 168시간만 제공합니다. 현재 데이터 품질표의 태양광 점검 대상 4개는 모델 잔차와
구분한 `데이터 패턴 신호`로 제공하며, 이를 센서 고장으로 단정하지 않습니다.

`serve-dashboard`는 `dashboard/`를 문서 루트로 지정하므로 어느 디렉터리에서 실행해도 같은 URL이
동작합니다. VS Code Live Server처럼 프로젝트 루트를 제공하는 기존 서버도 지원하도록 루트의
`solar_dashboard.html`과 과거 `map/html/solar_dashboard.html`에는 호환 이동 페이지를 둡니다.
루트 `forecast.html`, `model_analysis.html`과 과거 `plant_region_report_perm.html`도 같은 방식으로
현재 화면에 연결합니다.
`--output-dir`을 별도로 지정하면 HTML·CSS·JavaScript·JSON·GeoJSON을 함께 게시합니다.

루트의 빈 `xgboost/`와 `cnn_bilstm/`를 제품 구조로 사용하지 않습니다. 모든 모델 구현은
`src/solar_forecast/models/`, 새 모델·체크포인트는 `artifacts/models/`와
`artifacts/checkpoints/`에 둡니다. 2025년 구형 발전소별 CNN 체크포인트는 삭제하지 않고
`artifacts/legacy/cnn_bilstm/`로 이동했습니다. 자세한 판단은
[`docs/DATA_FORMAT_AND_DASHBOARD_AUDIT.md`](docs/DATA_FORMAT_AND_DASHBOARD_AUDIT.md)에 있습니다.

### 추가 발전소 후보

공개 범위를 넓히기 위해 한국농어촌공사 영암 2020~2025와 진도 2019, 남부발전 용수리·신풍리
최신 누적 파일도 실제로 내려받았습니다.
2022~2025는 영암1차·영암2차·율치 3개소, 105,191개 유효 시간 행으로 표준화할 수 있습니다.
2020~2021은 발전소 식별 컬럼이 없어 임의로 같은 설비로 간주하지 않고 격리합니다.

```powershell
python app.py audit-candidate-data
```

이 명령은 원본 hash, 스키마, 일계 합산, 중복·음수·결측, 발전소별 시간 커버리지와
Train/Validation/Calibration/Test 구간을 검사해 연도별 gzip 파티션과 후보 manifest를 만듭니다.
영암 데이터는 원문의 일계가 24개 1시간 버킷 합과 일치하고 세 설비 모두 시간 발전량이 공식
용량의 1.05배 이내임을 확인했습니다. 1시간 평균출력 `kW` 버킷은 에너지 수치상 `kWh`와 같으므로
MWh로 변환하고, 영암1·2차는 목포 ASOS 165, 율치는 강진군 ASOS 259라는 검토 매핑과 근거 URL을
버전 관리해 Gold에 자동 편입합니다. 한국지역난방공사
대구·신안 및 2026 인버터 실시간 API, 한국중부발전 발전량 API·169개소 안심구역 자료,
전력거래소 지역 집계의 용도와 우선순위는
[`docs/ADDITIONAL_DATA_SOURCE_AUDIT.md`](docs/ADDITIONAL_DATA_SOURCE_AUDIT.md)에 정리했습니다.

## 학습 컬럼 선정

기존 7개 기상·시간 변수만 유지하지 않고, 168시간 purge gap을 둔 3개 expanding
rolling-origin fold로 추가 컬럼을 검증했습니다. 뒤쪽 10% Calibration과 15% Test는 컬럼 선택에
사용하지 않았습니다. 기본 계약 `selected_v3_physics_aware_day_ahead`는 다음 26개 숫자 피처를 사용합니다.

- 기상: 기온, 강수, 일조, 일사, 풍속, 습도, 전운량, 중하층운량
- 시간: hour/dayofweek/month 및 시간·연중일 sin/cos
- 이력: 24시간 전, 168시간 전, 24시간 뒤로 민 7일 평균 발전량
- 정적: 설비용량, 설치각, ASOS 지점 위도·경도·해발고도
- 물리: 계산 태양고도, clear-sky 일사 proxy, 주야 여부

이력 피처는 예측 대상 시각보다 최소 24시간 이전 발전량만 사용합니다. 1·3·6시간 lag는 하루 전
예측 시점에 알 수 없으므로 금지했습니다. 증기압은 습도와 중복되고, 적설은 관측률이 1.12%,
최저운고는 관측률이 48.75%이며 MAE 개선도 없어 기본 피처에서 제외했습니다.

공식 표준 원본에서 타깃을 다시 집계하고 태양광 18개 발전소만 사용한 rolling-origin 재실험에서
23개 계약의 평균 MAE는 0.04028, 태양 위치 3개를 더한 26개 계약은 0.03904로 3.09%
개선됐습니다. 기상 관측 mask를 더한 30개 계약은 0.03937, 이력 가용성까지 더한 33개 계약은
0.03914로 26개 계약보다 나빠 기본 입력에는 넣지 않았습니다. 관련 컬럼은 품질 분석과 향후
결측 stress test용으로 Gold 데이터에 계속 보존합니다. 이 수치는 정식 Test 성능이 아니라
컬럼 선택용 Validation 평균입니다.

이 ablation은 legacy ASOS fail-closed 적용 전의 18개 안정 표본에서 얻은 기준선입니다. 현재
태양광 Gold는 22개소·701,011행을 보존하고, 시간 학습 적격은 21개소·671,731행입니다. 이 구성에서는
기존 결과를 고정 최적값으로 간주하지 않으며 `evaluate-features`와 모델별
Optuna를 다시 실행해 Validation 결과로 재선정합니다. 기존 수치는 탐색 시작점과 비교 기준으로만
사용하고 새 데이터에 하드코딩하지 않습니다. `evaluate-features`, XGBoost, CNN-BiLSTM은 모두
`quality_train_eligible` 컬럼을 필수 계약으로 요구하며, 컬럼이나 설정이 빠지면 필터 없이
진행하지 않고 즉시 중단합니다.

모델은 발전소별로 다른 피처 목록을 받지 않습니다. 모든 발전소가 위 26개 공통 컬럼을 같은 순서로
사용하고, 결측이어도 컬럼 자체는 유지합니다. 시퀀스는 `plant_id` 경계를 넘지 않으며, 현재 모델은
용량·설치각·관측소 좌표 같은 공통 정적 피처로 발전소 차이를 설명합니다. `plant_id` 자체는 현재
시퀀스 경계에만 쓰고 입력 embedding으로 넣지 않습니다. plant embedding은 기존 발전소 정확도와
신규 발전소 cold-start를 각각 group holdout으로 검증한 뒤 승격할 후보입니다.

## 결측치와 시간 분할 정책

결측은 실제 0과 같은 값으로 보지 않습니다. 발전량 음수와 시간 해상도가 맞지 않는 일 총량 자료는
원값과 원인 flag를 보존하되 학습에서 제외합니다. 용량초과는 단위·용량 메타데이터 문제와 실제
출력제어를 구분할 수 없어 자동 제외하지 않고 검토 flag로 남기며, 발전량 타깃 결측은 정답을
만들어 넣지 않습니다. 강수 공란은 다른
핵심 센서가 정상인 무강수 문맥에서만 0, 일조·일사 공란은 계산상 야간일 때만 0으로 해석합니다.
그 밖의 결측과 초기 168시간 history 결측은 그대로 유지합니다. 현재 전국 Gold는 history 결측을
이유로 행을 버리지 않으며 718,531행을 보존합니다.

- XGBoost: 중앙값으로 덮지 않고 `NaN`을 전달해 트리가 Train에서 결측 기본 분기 방향을
  학습합니다. 분할별 피처 결측률을 `preprocessing.json`에 남깁니다.
- CNN-BiLSTM: Train 행으로만 중앙값을 적합하고 원본 26개 피처마다 결측 mask 채널을 붙입니다.
  Train에서 전부 비어 있는 피처도 0+mask로 표현해 미래 통계가 역류하지 않습니다.
- CNN 창: 168시간 배열을 행마다 복제하지 않고 batch 시점에 자르는 lazy dataset을 사용합니다.
  서로 다른 발전소의 창은 절대 연결하지 않습니다.
- 메모리 상한: 회사×연도 gzip 파티션을 100,000행 chunk로 읽고 필요한 컬럼과
  `energy_source=solar`, `quality_train_eligible=true`를 먼저 적용합니다. 수치는 `float32`로
  축소하고 예상 프레임이 1,536MB를 넘으면 임의 샘플링 없이 `MemoryError`로 중단합니다.
- 실측: 27개 파티션 718,531행을 스캔해 태양광 Gold 701,011행 중 품질 게이트를 통과한
  671,731행·21개 발전소를 유지했습니다. XGBoost의 26개 피처와 문맥 컬럼을 포함한 DataFrame은
  114.68MB였습니다. 입력 gzip은 44,910,649 byte였고 1,536MB hard limit 아래에서 임의 샘플링
  없이 완료했습니다.
- 시간 분할: 모든 발전소에 같은 전역 날짜 경계를 적용합니다. Train 60%, Validation 15%,
  Calibration 10%, Test 15%이며 각 경계 앞에 최대 lookback 168시간을 비웁니다.
- 역할 분리: Validation은 조기 종료·피처·Hybrid gate 선택, Calibration은 잔차 임계값 고정,
  Test는 마지막 1회 성능 보고에만 사용합니다.

현재 평가는 매일 관측이 갱신되는 `rolling_origin_day_ahead_with_observation_updates`를 가정합니다.
따라서 Test 중 24/168시간 lag 사용은 허용됩니다. 연초 한 번에 1년 전체를 예측하는 fixed-origin
평가와는 다른 문제이므로 두 결과를 같은 지표로 섞지 않습니다.

동서발전의 하루 전 전국 태양광 예측 API가 제공하는 기온·습도·풍속·하늘상태·예측 일사량·설비용량은
운영 예보용 후보로 출처 카탈로그에 등록했습니다. 서비스 키가 필요한 선택 수집원이며, API 자체의
`태양광 발전량 예측값`은 현재 모델의 기본 입력에서 제외하고 외부 기준모델로만 별도 평가합니다.

```text
manifest.json
report.html
model/<실행시간>/cnn_bilstm.pt
```

저장 수준은 다음과 같이 선택할 수 있습니다.

```bash
# 최종 모델·지표·HTML만 저장(기본)
python app.py pipeline target --data data.csv --artifact-level minimal

# 이상징후 상세를 gzip CSV로 추가 저장
python app.py pipeline target --data data.csv --artifact-level standard

# 전처리 CSV까지 저장하여 문제를 추적
python app.py pipeline target --data data.csv --artifact-level debug
```

실행 중 오류가 발생하면 해당 실행 폴더에만 `error.log`가 생성됩니다. 파일에는 실패 단계, 입력 경로, 예외 유형·메시지와 traceback이 기록되며, 프로세스는 실패 종료 코드를 유지합니다. 정상 실행에서는 에러 파일을 생성하지 않습니다.

## 이상징후 해석 범위

탐지 결과는 `기상 영향`, `예상 대비 저발전`, `예상 대비 과발전`, `급격한 출력 변화`, `발전 정지 의심`, `데이터 품질 문제`, `미확인 외부요인`으로만 표현합니다. `발전 정지 의심`은 고장 판정이 아니며 보고서와 알림에는 다음 한계를 함께 표시합니다.

> 공개 데이터만으로 설비 고장·정비·출력제어 여부를 확인할 수 없습니다.

동일 지역의 여러 발전소가 동시에 벗어나는 경우 기상 영향 가능성을 높이고, 한 발전소만 벗어나면 `발전소 개별 미확인 요인`으로 분류합니다. 어느 경우에도 인버터 또는 설비 고장을 확정하지 않습니다.

이상징후 임계값은 평가할 Test 잔차의 상위 일정 비율을 직접 고르지 않습니다. 독립 Calibration 절대잔차로
임계값을 한 번 고정한 뒤 Test에 적용하므로 Test 이상치 개수는 1%나 5%로 강제되지 않습니다.
`--contamination`은 Calibration 임계 분위수의 목표값이며 Test 행을 순위화하는 값이 아닙니다.

제품 코드는 모두 `src/solar_forecast/` 아래에 있습니다. `ForecastPipeline`은 데이터 저장소,
전처리기, 학습 어댑터, 보고서 작성기를 생성자에서 주입받으며, `CollectionService`와
`HybridExperiment`도 수집 및 앙상블 실행 경계를 각각 소유합니다. 자세한 의존성 방향은
`docs/ARCHITECTURE.md`를 참고합니다.

대량 데이터 취업 포트폴리오 관점의 Bronze/Silver/Gold 경계, 파티셔닝, hash lineage, atomic write,
quarantine, lazy sequence와 Parquet/Polars/Spark 확장 기준은
[`docs/SCALABLE_DATA_ENGINEERING.md`](docs/SCALABLE_DATA_ENGINEERING.md)에 별도로 정리했습니다.

## 독립 모델 학습

```bash
python app.py train xgboost
python app.py train cnn_bilstm
python app.py status
```

모델 설정은 `config/models/`에 있으며 현재 보유 파일의 컬럼명과 일치하도록
`file/standardized/model_ready_parts/`, `generation_mwh`, `selected_v3_physics_aware_day_ahead`의
26개 피처와 `energy_source=solar`, `quality_train_eligible=true` 필터가 명시되어 있습니다.
처음 실행하거나 원본을 갱신한 뒤에는 `python app.py prepare-data`를 먼저 실행해야 합니다.
학습 job은 모든 파티션의 발전소에 같은 전역 시간 경계를 적용해 네 역할을 내부에서 나눕니다.

두 독립 모델은 기본적으로 Optuna를 사용합니다. XGBoost는 최대 30 trial/2시간, CNN-BiLSTM은
최대 20 trial/4시간이며 `artifacts/optimization/solar_models.db`의 모델별 study를 재개합니다.
실제 study 이름에는 데이터·학습 설정 fingerprint가 자동으로 붙으므로 원본 또는 의미 있는 설정이
바뀌면 동일한 표시 이름을 사용해도 별도 탐색으로 시작합니다.
`max_trials`는 실행할 때마다 더하는 수가 아니라 해당 study의 **최대 누적 trial 수**입니다.
현재 고정 파라미터는 과거 데이터 구성에서 얻은 Optuna 결과로서 문서상의 참고 기준일 뿐입니다.
새 데이터·피처 구성의 study에는 enqueue하거나 우선권을 주지 않으며 독립적으로 다시 탐색합니다.

학습 상태는 기본적으로 `artifacts/checkpoints/<model>/<data-config-fingerprint>/`에 원자적으로
저장합니다. CNN-BiLSTM은 매 epoch마다 모델·optimizer·early stopping·최적 모델·난수 상태를,
XGBoost는 50 boosting round마다 Booster를 저장합니다. Optuna trial도 같은 단위로 이어서
실행하며 프로세스가 비정상 종료된 trial은 heartbeat로 판별해 같은 파라미터로 한 번 재시도합니다.
데이터 manifest 또는 학습 설정이 바뀌면 fingerprint가 달라져 과거 상태를 잘못 읽지 않습니다.
완료 상태도 최종 manifest 저장 전 장애에 대비해 유지하므로 같은 입력으로 다시 실행하면 이미
끝난 최종 학습을 반복하지 않습니다. 새로 학습하되 상태는 계속 저장하려면 모델 설정의
`checkpoint.resume`을 `false`로 설정합니다. 상세 계약과 검수 결과는
[`docs/CHECKPOINT_AND_CONSISTENCY_AUDIT.md`](docs/CHECKPOINT_AND_CONSISTENCY_AUDIT.md)를 참고합니다.

- 탐색 입력: Train과 Validation만 사용
- 목적함수: Validation MAE
- XGBoost: boosting iteration별 pruning + early stopping
- CNN-BiLSTM: epoch별 pruning + early stopping, lazy sequence 중 제한된 대표 표본으로 탐색
- 최종 학습: 선택된 설정으로 전체 Train을 다시 사용
- Calibration: 잔차 임계값·구간 보정용으로 예약
- Test: 설정 선택이 끝난 뒤 마지막 성능 보고에만 1회 사용

시간을 더 줄이거나 고정 설정을 재현할 때만 CLI에서 제한을 덮어씁니다.

```bash
python app.py train xgboost --max-trials 10 --optimizer-timeout-seconds 1800
python app.py train cnn_bilstm --max-trials 5 --optimizer-timeout-seconds 3600
python app.py train xgboost --no-optuna
python app.py train cnn_bilstm --smoke
```

각 실행 폴더에는 `optimization_summary.json`, `optimization_trials.csv`, 최종 `best_params.json`과
모델 manifest가 저장됩니다. 두 모델 모두 Validation/Calibration/Test 예측에
`timestamp, plant_id, region, plant, y_true, y_pred` 공통 문맥을 보존하므로 대시보드가 전체 파일을
RAM에 올리지 않고 임시 SQLite에서 같은 Test 행만 정렬할 수 있습니다. 데이터·피처·분할 또는
탐색공간을 바꿀 때는 기존 trial을 섞지 않도록 설정의 `study_name` 버전을 올립니다. `--smoke`는
배선 검사 목적이므로 Optuna를 자동 생략하며 완료 manifest에도 `execution_mode=smoke`를 보존해
정식 성능 화면에서 자동 제외합니다.

## 포트폴리오 실험 구분

실험 정의는 `config/experiments/` 아래에 분리되어 있습니다.

- `controlled.json`: 동일 lookback, 동일 피처 정보, 동일 시간 분할의 구조 비교
- `optimized.json`: 같은 태양광 표본·피처·분할에서 Validation-only Optuna를 적용한 성능 비교
- `hybrid.json`: Validation의 지역·시간대별 모델 신뢰도를 근거로 매 행의 결합비를 결정하는 동적 Hybrid

Hybrid 입력 CSV는 두 모델의 예측이 같은 행에 정렬되어 있어야 합니다.

```text
timestamp,region,plant,y_true,xgb_pred,cnn_pred
```

```bash
python app.py hybrid output/validation_predictions.csv output/test_predictions.csv \
  --output-dir output/experiments/hybrid --artifact-level minimal
```

원본 병합 파일의 `일시`, `지역`, `발전구분`, `합산발전량(MWh)`는 각각
`timestamp`, `region`, `plant`, `y_true`로 자동 매핑됩니다. 다만
`xgb_pred`와 `cnn_pred`는 동일 Validation/Test 행에 대해 각 기본 모델이 만든 예측값이어야
하며 원본 관측 파일만으로 임의 생성하지 않습니다. 키
`timestamp + region + plant`는 중복 없이 두 모델 사이에서 정확히 정렬되어야 합니다.

Hybrid는 고정 가중치가 아닙니다. Validation에서 발전소·지역·시간대·일출/주간/일몰 regime와
두 모델의 예측 불일치 구간별 최적 결합비와 MAE를 계산합니다. Test의 각 행은 가장 구체적이면서
표본이 충분한 근거를 사용하고, 부족하면 발전소/지역/전체 Validation 순으로 후퇴합니다.
Test의 `y_true`는 평가에만 사용하며 결합비 결정에는 사용하지 않습니다.

`--artifact-level standard` 이상에서 저장되는 `hybrid_predictions.csv.gz`에는
`gate_scope`, `xgb_expected_mae`, `cnn_expected_mae`, `selected_model`,
`model_disagreement`, `decision_reason`, 두 모델의 실제 결합비가 함께 기록됩니다.
`dynamic_gate_profiles.csv`에는 판단에 사용한 Validation 근거가 저장됩니다.

## 테스트

```bash
.\.venv\Scripts\python.exe -m pytest -q
```

후보 컬럼 비교는 Calibration/Test를 사용하지 않는 3-fold purged rolling-origin으로 재현할 수 있습니다.

```bash
python app.py evaluate-features --folds 3 --validation-window-hours 2160 --gap-hours 168
```

논문 근거, 발표자료 방식의 문제점, 발전소별 표준화와 모델 전환 순서는
[`docs/MODEL_PREPROCESSING_ROADMAP.md`](docs/MODEL_PREPROCESSING_ROADMAP.md)에 정리했습니다.

## 데이터 출처

- 기상청 기상자료개방포털 ASOS 시간자료
- 한국남동발전·한국남부발전·한국동서발전·한국서부발전 공개 태양광 발전량 및 메타데이터
- 한국농어촌공사 영암 태양광 2020~2025 주기성 과거파일(후보 staging; 2022~2025만 개체 식별 가능)
- 국가데이터처 SGIS 행정구역 통계 및 경계(2025년 2분기 시도 경계)

발전사별 원본 기준일이 다를 수 있으므로 모델 비교용 데이터셋은 네 출처가 모두 존재하는 공통 기간으로 고정합니다. 운영 예측에는 관측 ASOS와 별도로 예측 발행시각이 보존된 기상예보 데이터가 필요합니다.
