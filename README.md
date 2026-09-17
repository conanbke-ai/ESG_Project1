# 전국 재생에너지 표준화·태양광 발전량 예측·이상징후 알림 시스템

현재 실측 기반 모델 최적화·하이브리드 채택 작업은 [실행 및 평가 계약](docs/OPTIMIZED_MODEL_BENCHMARK.md)을 따릅니다. `python app.py benchmark --plan`으로 모델별 탐색과 1·24·72시간 비교 설정을 확인할 수 있습니다. 과거 문서의 고정 24시간·동일 특징 조건보다 이 계약이 우선합니다.

본 학습은 사용자 Windows 로컬 GPU에서 수행합니다. ChatGPT 작업 환경과 CI에서 본 학습을 시작하지 않습니다. 로컬에서 `python tools/run_local_benchmark.py --data file/standardized/model_ready_parts`를 사용하며 CUDA GPU가 없으면 학습 전에 종료합니다. 기본 날짜 고정 후보와 탐색 예산을 유지하고, `--preflight-only`는 학습 전 점검만 수행합니다.

2026-09-17 전처리 보완은 발전량 원값·ASOS 결측 사유를 보존하고 CNN의 Train 전용 입력
표준화 통계를 저장합니다. 확대 자료의 날짜 고정 분할 후보와 학습 없는 커버리지 점검은
[분할 점검 방법](docs/OPTIMIZED_MODEL_BENCHMARK.md#날짜-고정-후보와-학습-전-점검)을 따릅니다.

**서비스 표시 모집단과 AI 모델 모집단은 의도적으로 분리합니다.** 전국 발전설비 현황 화면은
AI 학습 적격 여부로 발전소를 제거하지 않고 EPSIS의 공식 공개 전국 등록 범위를 유지합니다.
반면 XGBoost·CNN-BiLSTM·Hybrid 학습/평가는 발전원, 시간 해상도, 식별·기상 매핑 근거,
시계열 연속성, temporal split 및 purge 조건을 통과한 별도 모델 적격 subset만 사용합니다.
학습 제외는 원본 삭제가 아니라 `model eligibility` 상태입니다. 설계 의도와 과거/현재 분할 기준은
[서비스 전체 모집단과 AI 학습 모집단 분리 정책](docs/DATA_POPULATION_AND_MODEL_ELIGIBILITY.md)에 정리합니다.

CI를 사용하지 않는 학습 준비는 `python tools/audit_forecast_readiness.py`로 실제 모델별
표본·연속 입력창을 먼저 확인합니다. [로컬 학습·재개·저장 모델 검증 절차](docs/OPTIMIZED_MODEL_BENCHMARK.md#ci-없이-로컬에서-실행)를 제공합니다.

[실측 실행 결과](docs/OBSERVED_BENCHMARK_RESULT.md)에는 공식 관측자료로 수행한 1·24시간 제한 실험과 저장 모델 재검증 근거를 기록했습니다. 24시간 Hybrid의 선택 구간 개선은 최종 Test에서 유지되지 않았습니다. 전체 데이터 최적화와 과거 저장 모델의 성능 재현은 아직 완료하지 않았습니다.

후속 검수에서 실행 환경을 바꾼 CNN 재학습 결과의 차이도 확인했습니다. CPU 검증 환경의 패키지 버전을 고정하고, 저장 모델 재예측과 별도로 새 학습을 반복 비교하는 CI를 추가했습니다. 같은 환경의 재현성과 서로 다른 환경의 결과 비교를 구분합니다.

## 코드 구조와 ASOS API 설정

- 전체 폴더·파일·함수 생성 규칙: [PROJECT_STRUCTURE.md](docs/PROJECT_STRUCTURE.md)
- Python 파일별 목적과 전체 함수 색인: [CODE_INDEX.md](docs/CODE_INDEX.md)
- 변경 전후 모델 수치 비교와 실데이터 성능 검증 범위: [MODEL_PARITY_VALIDATION.md](docs/MODEL_PARITY_VALIDATION.md)
- 서비스 전체 모집단과 AI 학습 모집단 분리: [DATA_POPULATION_AND_MODEL_ELIGIBILITY.md](docs/DATA_POPULATION_AND_MODEL_ELIGIBILITY.md)
- 발급받은 ASOS API 키 적용: [KMA_ASOS_API.md](docs/KMA_ASOS_API.md)

실제 ASOS 키는 프로젝트 루트 `.env.local`의 `KMA_ASOS_SERVICE_KEY`에 일반인증키(Decoding)로 설정합니다. `.env.example`에는 빈 예시만 둡니다. 기존 `DATA_GO_SERVICE_KEY`는 중부발전용입니다.

```powershell
python app.py collect --sources kma --kma-mode api --station-ids 108,159 --start-date 2026-01-01 --end-date 2026-01-02 --api-max-calls 10
python app.py prepare-data
```

모델 구현은 `models/cnn_bilstm/`, `models/xgboost/`, `models/hybrid/`로 모았습니다. CNN의 실제 네트워크는 `network.py`, 학습 설정 연결은 `trainer.py`, 학습 실행은 `training_workflow.py`, 저장 모델 평가는 `evaluation.py`에서 찾습니다. 대시보드 원본은 `dashboard/src/`, 생성된 배포 파일은 `dashboard/assets/dashboard.js`입니다.

```powershell
python tools/build_dashboard_assets.py
python tools/update_code_index.py
python tools/check_structure.py
```


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

각 실행 결과는 기본적으로 `artifacts/pipeline/<실행시간>/`에 저장됩니다. 단계 사이의 DataFrame은 메모리로 전달하고 기본 `minimal` 모드에서는 최종 산출물만 저장합니다.

## 공식 데이터 자동 수집

현재 자동화된 4개 발전사 파일은 각 공식 홈페이지의 공공데이터 게시물 또는 그 게시물이 연결한
공공데이터포털 첨부파일에서 내려받습니다. 한국농어촌공사 영암 원본은 확보된 주기성 파일을
`prepare-data`가 자동 심사·편입합니다. `collect`가 새로 만든 발전사 Silver 파일도
plant-hour 발전량 계약을 만족하면 registry/admission을 거쳐 Gold 후보에 자동 연결됩니다.
발전사별 다운로드 URL 환경변수는 필요하지 않습니다.
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
파티션으로 변환합니다. 이어 `file/standardized/downloads/`의 collector Silver를 검사해
plant-hour 계약을 만족한 3개 파일·195,225행을 registry 후보로 넘기고, 지역 단위 학습파일처럼
발전소 식별자가 없는 1개 파일은 `collector_admission_manifest.json`에 사유를 남겨 제외합니다.
농어촌공사 영암 6개 파일은 개체 식별이 가능한 2022~2025 4개 파일·105,191개 시간 행을 같은
계약으로 편입하고, 2020~2021은 식별자 부재로 격리합니다. 기본 원본 검사는
`generation_manifest.json`, collector 검사는 `collector_admission_manifest.json`, 영암 검사는
`candidates/krc_yeongam/candidate_manifest.json`에 행 수·기간·결측·음수·중복·물리 상한과 함께 기록합니다.
각 파티션은 한 원본 파일 단위로 처리하고 완성된 임시 파일만 원자적으로 교체합니다. manifest에는
입력·출력 byte와 SHA-256을 기록해 같은 파일명의 수정본과 재처리 lineage를 추적합니다.
KOSPO의 `신재생사업본부`는 개별 설비명이 아니므로 공식 발전기명 `부산 철도태양광 #2`가 함께
확인된 행만 사양표의 `부산철도 2`로 식별합니다. 신규 수집과 보관된 collector Silver의
admission/registry/Gold 읽기에 같은 규칙을 적용하고, 원본 파일은 변경하지 않습니다.
주소·기상 매핑과 용량의 검증은 별도입니다. 해당 설비의 용량은 현재 사양표에서 비어 있고
부산진구 주소에 대한 기상 매핑도 미승인 상태이므로, 식별자 보정만으로 Gold에 편입하지 않습니다.
검증 범위와 남은 게이트는 [`docs/KOSPO_IDENTITY_AUDIT.md`](docs/KOSPO_IDENTITY_AUDIT.md)에 기록합니다.
이후 공식 표준 발전량을 발전소·시간 단위로 재집계하고, 기존 병합본의 ASOS 지점번호는 승인
근거가 아닌 audit-only 후보로만 보존합니다. 공식 주소·좌표, 발전기간 전체를 덮는 KMA 지점 이력,
근거가 있는 reviewed mapping을 통과한 행으로 단일 호환본 `file/standardized/model_ready.csv.gz`와
학습용 회사×연도 파티션 `file/standardized/model_ready_parts/`를 생성합니다. 현재 표준 후보는
historical/KRC/collector를 합쳐 4,536,981행이고, registry/ASOS 근거 통과 후 867,397행,
누적 snapshot 최신본 선택 후 721,363행으로 좁혀집니다. 2026년 기상 연간 파일이 아직 없어
해당 발전량 2,832행·2개소는 Silver에는 보존하고 Gold에서만 보류하므로 최종 Gold는
718,531행·24개 발전소(태양광 22, 풍력 2)입니다. 태양광 44개와 소수력 1개를 포함한
45개 자산은 근거 없는 기상 결합 대신 registry에 격리합니다.
태양광 Gold 701,011행·22개소 가운데 여수태양광 29,280행은 일 총량이 고정 야간 버킷에 적재된
자료라 원문과 품질 근거는 보존하되 시간 예측 학습에서는 제외합니다. 따라서 현재 실제 태양광
학습 적격 범위는 671,731행·21개소입니다.
기존 병합 타깃과 공식 원본의
차이는 `legacy_pipeline_quality_report.csv`에 기록하고, 검증된 공식 타깃만 학습에 사용합니다.

## Dashboard

```powershell
python app.py build-dashboard
python app.py serve-dashboard --port 5500
```

`solar_dashboard.html`은 모델 학습 적격성으로 필터링하지 않은 EPSIS 전국 등록 범위를 사용합니다.
`forecast.html`과 `model_analysis.html`은 모델의 정식 평가 모집단과 결과를 사용합니다. 즉, 기본 화면의
전체 발전소 현황과 AI 학습·예측 가능 발전소 수는 서로 다른 지표이며 같아야 할 필요가 없습니다.

## License

프로젝트에서 사용하는 외부 데이터·문헌은 각 원 출처의 이용 조건을 따릅니다.
