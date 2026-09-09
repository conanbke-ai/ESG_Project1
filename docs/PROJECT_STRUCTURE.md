# 디렉터리·파일·함수 생성 규칙

제품은 하나의 Python 패키지와 정적 대시보드로 유지한다. 모델 이름, 파일 역할, 실행 결과 위치를 구분하고 새 구현은 아래 소유 영역에 둔다. 모든 Python 파일의 실제 클래스·메서드·내부 함수는 [코드 색인](CODE_INDEX.md)에서 찾을 수 있다.

## 프로젝트 루트에서 찾는 법

| 위치 | 무엇을 위한 것인가 | 생성·수정 규칙 |
| --- | --- | --- |
| `app.py` | 개발용 CLI 진입점 | 명령 dispatch만 한다. 업무 로직을 넣지 않는다. |
| `src/solar_forecast/` | 실행하는 Python 제품 코드 | 아래 기능별 소유 영역에 배치한다. 데이터·가중치를 저장하지 않는다. |
| `config/models/<model_id>.json` | 모델별 학습 조건 | 구현 폴더와 같은 모델 ID를 쓴다. 비밀키를 넣지 않는다. |
| `config/experiments/` | 비교 실험·후보 연구 조건 | 구현 완료 여부와 실험 정의를 구분한다. 후보 이름만으로 구현된 모델이라고 표시하지 않는다. |
| `config/architecture/` | 파일 역할과 이전 경로 대응표 | 새 모듈은 `modules.json`에 목적을 등록한다. |
| `config/*.json` | 검토한 발전소·기상 연결·행정구역·수신 경로 계약 | 설정의 근거와 버전을 유지한다. 실제 연락처는 로컬 설정에 둔다. |
| `file/solar_data_file/` | 기존 공식 발전량 원본 보관 | 원본 이름과 byte를 보존한다. Python 파일을 만들지 않는다. |
| `file/raw/` | 새로 수집한 Bronze와 수집 실행 manifest | 공급자 하위에 원본을 보관한다. ASOS는 관측소·기간·응답 snapshot으로 나눈다. |
| `file/KMA_data_file/` | 관측소 메타데이터와 연도별 기상 자료 | `META_관측지점정보.csv`, `OBS_ASOS_TIM_<year>.csv` 계약을 유지한다. |
| `file/standardized/downloads/` | collector가 만든 발전량 Silver | admission이 심사한다. 파일 수집 성공만으로 Gold가 되지 않는다. |
| `file/standardized/generation/` | 보관 원본에서 생성한 표준 발전량 | 기관·연도 파티션과 manifest로 관리한다. |
| `file/standardized/model_ready_parts/` | 모델이 읽는 Gold 파티션 | `model_ready_manifest.json`의 schema·lineage를 따른다. |
| `file/merge_data/` | 이전 병합 입력의 호환 보관소 | 새로운 업무 코드를 추가하지 않는다. 기존 설정이 참조하므로 경로를 임의 변경하지 않는다. |
| `data/` | 공식 전국 설비 원장 등 별도 근거 데이터 | `config/national_solar_inventory.json`의 출처·hash 계약을 따른다. |
| `map/` | 공식 지도 원본·경계와 과거 화면 진입점 | 지도 변환은 `reporting/sgis_boundaries.py`가 맡는다. |
| `artifacts/models/<model_id>/<run_id>/` | 학습 완료 모델·예측·실행 manifest | 모델 구현 코드와 분리한다. 기존 run ID·파일 형식을 유지한다. |
| `artifacts/checkpoints/` | 중단 학습 재개 상태 | config/dataset fingerprint로 호환성을 확인한다. |
| `artifacts/optimization/` | Optuna study 저장소 | 모델별 study ID를 사용한다. |
| `artifacts/pipeline/` | 전체 파이프라인 실행 결과 | 이전 기본값 `output/pipeline/` 대신 새 실행부터 사용한다. |
| `artifacts/experiments/hybrid/` | 하이브리드 결합 실험 결과 | 기존 `output/experiments/hybrid/` 대신 사용한다. |
| `artifacts/evaluation/features/` | 특징 비교 평가 결과 | 기존 `output/evaluation/features/` 대신 사용한다. |
| `artifacts/verification/`, `artifacts/notifications/` | 연결 검증 보고서·알림 outbox | 각 job manifest/DB 계약을 유지한다. |
| `artifacts/legacy/`, `output/` | 이전 실행 결과 | 기존 파일을 임의 삭제하거나 새 모델 소스로 취급하지 않는다. |
| `dashboard/src/` | 사람이 수정하는 화면별 JavaScript 원본 | `context`, `formatting`, `coverage`, `forecast`, `analysis`, `bootstrap`의 역할을 지킨다. |
| `dashboard/assets/dashboard.js` | 배포용으로 생성한 JavaScript | 직접 수정하지 않는다. `tools/build_dashboard_assets.py`로 만든다. |
| `dashboard/assets/dashboard.css`, `dashboard/*.html` | 화면 스타일과 실제 정적 화면 진입점 | 화면 내용·스타일 변경은 이곳에서 한다. |
| `dashboard/data/` | 화면이 읽는 생성 JSON·GeoJSON | Python reporting 결과다. 모델 구현이 아니다. |
| 루트의 `*.html`, `map/html/solar_dashboard.html` | 예전 URL을 현재 화면으로 연결하는 이동 페이지 | 기능을 추가하지 않는다. 이전 북마크 호환을 위해 유지한다. |
| `tests/` | 자동 회귀 검증 | `test_<대상 기능>.py`, `test_<조건과 기대 결과>`로 이름 짓는다. |
| `tools/` | 구조 검사·색인·정적 파일 생성 도구 | 제품 CLI 업무 로직을 넣지 않는다. |
| `docs/` | 설계·운영·데이터 근거 문서 | 현재 구조는 이 문서, 실제 심볼은 코드 색인을 기준으로 한다. |

원본과 기존 체크포인트를 한 번에 이동하면 입력 경로·manifest·재개 fingerprint가 달라질 수 있다. 이번 변경은 Python/JavaScript 구현을 실제로 재배치하고 새 평가 산출물 위치를 통일한다. 보관 데이터의 물리 경로와 기존 모델 설정 JSON 값은 유지한다. Windows 작업 폴더의 잠긴 빈 디렉터리를 원격에서 삭제한 것으로 간주하지 않는다.

## Python 코드 소유 영역

| 폴더 | 책임 | 대표 파일 |
| --- | --- | --- |
| `cli/` | 명령 옵션과 인자를 업무 객체로 전달 | `parser.py`, `collection_commands.py`, `training_commands.py` |
| `collectors/` | 공식 공급자에서 원본 수집·공급자 스키마 변환 | `kma_api.py`, `kma_api_client.py`, `generation_collectors.py`, `generation_normalizers.py` |
| `datasets/` | admission·registry·저장소·Gold 생성 | `collector_admission.py`, `plant_registry.py`, `model_dataset_builder.py`, `preparation_service.py` |
| `features/` | 관측값과 과거 데이터에서 학습 특징 계산 | `asos_features.py`, `history_features.py` |
| `quality/` | 데이터의 물리·통계 품질과 학습 적격 여부 | `generation_quality.py` |
| `models/` | 모델별 학습·평가 구현 | 아래 모델 표 참조 |
| `evaluation/` | 공통 시간 분할·회귀 지표·특징·예측 동등성 비교 | `temporal_split.py`, `regression_metrics.py`, `feature_ablation.py`, `model_parity.py` |
| `jobs/` | 독립 실행·잠금·manifest·dispatcher orchestration | `training_job.py`, `verification_job.py`, `notification_dispatch_job.py`, `contracts.py` |
| `pipeline/` | 데이터·학습·리포트 adapter를 잇는 전체 실행 | `forecast_pipeline.py`, `cnn_training_adapter.py`, `contracts.py` |
| `anomalies/` | 운영 이상치 이벤트와 설명 정책 | `event_batch.py`, `influence_policy.py` |
| `notifications/` | outbox·수신 경로·외부 알림 전송 | `outbox_repository.py`, `recipient_routes.py`, `delivery_provider.py` |
| `reporting/` | 모델·품질·설비 자료를 화면/보고서로 투영 | `dashboard_builder.py`, `model_analytics.py`, `national_solar_inventory.py` |
| `infrastructure/` | 환경·파일 저장·오류·기본 경로 | `environment.py`, `artifact_store.py`, `project_paths.py` |

CLI 도움말과 job 조회는 학습 프레임워크를 import하지 않는다. 패키지 `__init__.py`는 공개 export 경계이며 수집·학습을 시작하지 않는다. 기능 구현이 CLI 처리 함수를 역으로 import하지 않도록 구조 검사에서 막는다.

## 모델 폴더를 읽는 법

| 모델 | 위치와 내부 파일 | 의미 |
| --- | --- | --- |
| CNN-BiLSTM | `models/cnn_bilstm/network.py` | 실제 CNN·BiLSTM 층과 forward, `CnnBiLstmNetworkConfig` |
| CNN-BiLSTM | `sequence_config.py`, `sequence_data.py` | 입력 시계열 길이·분할·batch 설정과 lazy window dataset |
| CNN-BiLSTM | `trainer.py` | 공통 모델 JSON을 읽어 학습 workflow를 호출하는 adapter |
| CNN-BiLSTM | `training_workflow.py`, `optimization.py` | 학습 실행·체크포인트·예측 저장과 epoch/Optuna 최적화 |
| CNN-BiLSTM | `evaluation.py`, `adaptive_training.py` | 저장 모델 평가·이상치 분석, 선택적 적응 학습 |
| XGBoost | `models/xgboost/trainer.py` | XGBoost 데이터 준비·학습·모델 파일·예측 기록 |
| XGBoost | `optimization.py`, `checkpoint.py` | 전용 하이퍼파라미터 탐색과 boosting 중단/재개 |
| Hybrid | `models/hybrid/dynamic_gate.py`, `experiment.py` | Validation 근거로 두 모델을 결합하고 실험 결과 생성 |
| 공통 지원 | `models/shared/checkpoint_store.py`, `optuna_study.py` | 모델 공통 fingerprint·상태 저장·탐색 저장소. 별도 예측 모델이 아니다. |

기존 `cnn/` 엔진과 바깥 `cnn_bilstm.py` adapter는 `cnn_bilstm/` 한 폴더에 모였다. XGBoost의 바깥 최적화·체크포인트 파일도 `xgboost/`로 옮겼다. 기존 모델 ID `cnn_bilstm`, `xgboost`, `hybrid`와 `.pt`의 state_dict/config payload는 유지한다. 내부 import 경로 변경은 [module_moves.json](../config/architecture/module_moves.json)에 기록한다. 오래된 내부 import를 사용하는 별도 스크립트는 이 대응표에 맞춰 수정한다.

## 이름과 새 파일 생성 규칙

1. Python 파일·함수·변수는 `snake_case`, 클래스는 `PascalCase`, 상수는 `UPPER_SNAKE_CASE`를 쓴다. 수치에는 필요한 단위를 붙인다: `generation_mwh`, `capacity_mw`, `timeout_seconds`.
2. 파일명은 대상과 역할을 드러낸다: `kma_api_client.py`, `plant_registry.py`, `training_job.py`. 독립 파일명 `service.py`, `config.py`, `models.py`, `base.py`, `utils.py`, `helpers.py`, `model.py`, `data.py`는 만들지 않는다. 폴더 문맥이 명확한 `trainer.py`, `network.py`, `contracts.py`는 허용한다.
3. 함수는 동작과 대상을 표현한다: `parse_asos_page`, `validate_asos_observations`, `merge_asos_observations`, `train_cnn_bilstm_epoch`. 표준 인터페이스 `fit`, `predict`, `transform`, `run`, `collect`, PyTorch `forward`는 클래스 문맥과 호환성을 유지한다.
4. 내부 helper는 `_`로 시작한다. 여러 영역에서 필요해지면 목적 있는 공통 모듈로 옮긴다. 동일 함수를 여러 파일에 복사하거나 `new`, `final`, `v2`, `backup`을 붙여 코드 복제본을 만들지 않는다.
5. JavaScript 함수는 `camelCase`를 쓰며 화면 문자열 생성은 `render…`, 이벤트 연결은 `bind…`, 표시 값 변환은 `format…`로 드러낸다. 모델 계산·데이터 품질 판단을 화면 코드에 복제하지 않는다.
6. 새 모듈을 만들기 전에 소유 폴더를 정하고 모듈 docstring과 `config/architecture/modules.json`의 목적을 함께 작성한다. 새 공개 함수는 입력·반환값, 단위, 파일/네트워크 부작용을 설명한다.
7. 입출력 계약이 바뀌면 `jobs/contracts.py` 또는 해당 데이터 계약의 버전과 소비자를 함께 수정한다. 단순 코드 이동으로 학습 설정·dataset fingerprint·시간 분할·지표 정의를 바꾸지 않는다.
8. 기본 저장 위치는 `infrastructure/project_paths.py`에서 찾는다. 모델 산출물은 모델 ID와 run ID, 데이터는 공급자·기간·파티션, 수집 응답은 source와 request 범위로 구분한다. API 키는 파일명·manifest·오류 URL에 기록하지 않는다.

## 변경 후 실행

```powershell
python tools/build_dashboard_assets.py
python tools/update_code_index.py
python tools/check_structure.py
python -m compileall -q src tests tools
python -m pytest
```

구조 검사는 모듈 목적 누락, 모호한 파일명, 깨진 내부 import/심볼, 소스 폴더에 들어온 데이터·가중치, 생성 파일과 원본 불일치를 실패로 처리한다. 계산이나 저장 계약을 바꾼 경우 해당 회귀 검증도 통과해야 한다. 전체 의존성이 없는 환경의 구조 검사 성공을 전체 모델 학습 검증으로 표시하지 않는다.

실제 변경 전후 수치 비교는 [MODEL_PARITY_VALIDATION.md](MODEL_PARITY_VALIDATION.md)를 따른다. `tools/verify_model_reorganization.py`는 별도 checkout의 비교 실험을 실행하고, 제품 CLI의 `compare-predictions`는 이미 저장한 예측 CSV를 비교한다.
