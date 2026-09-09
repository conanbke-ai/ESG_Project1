# 실측 자료 기반 모델 최적화와 비교

목적은 취업 포트폴리오에서 공식 발전량·기상청 ASOS 관측자료를 처리하고, 모델별 최적화 과정과 하이브리드 채택 근거를 재현 가능한 화면으로 보여주는 것입니다. 예보 API 가입은 필요하지 않습니다. 기본 1·24·72시간은 비교 과제의 예시이며 24시간으로 고정하지 않습니다. CNN-BiLSTM이 장기 예측에 반드시 우수하다고 가정하지 않습니다.

## 실행

```powershell
python app.py prepare-data --no-collected-downloads
python app.py benchmark --plan
python app.py benchmark --config config/experiments/optimized.json
python app.py serve-dashboard
```

`benchmark`는 완료 후 대시보드 데이터를 갱신합니다. `--plan`은 실제 후보·피처·탐색 예산을 출력하며 학습하지 않습니다. `--smoke`는 1시간·24시간 입력창·소량 데이터로 연결만 점검하고 대시보드의 성능 비교에서 제외합니다.

## 평가 과제

목표 발전량 시각이 `t`, 예측 기간이 `h`이면 예측 시작 시각은 `t−h`입니다. 각 모델은 `h`시간 뒤 한 시점의 시간당 발전량을 예측합니다. 화면의 시계열은 시작 시각을 이동하며 만든 예측들을 연결한 것입니다. XGBoost는 `t−h`의 관측 피처, CNN은 `t−h`까지 포함하는 연속 시간창을 사용합니다. 예측 CSV에는 목표시각, 시작시각, 기간, 발전소 ID, 실측값, 예측값과 시작시각 발전량을 유지한 기준선이 저장됩니다. 미래에 관측된 날씨를 과거 예측 입력으로 사용하지 않습니다.

이는 이미 확보한 과거 관측자료로 수행하는 rolling-origin 평가입니다. 시각 `t−h`의 시간 관측이 그 시각에 사용 가능하다고 가정하며, 실제 API 발행 지연은 재현하지 않습니다. Test 진행 중 새로 관측된 이전 시점 자료를 입력으로 갱신할 수 있지만 모델 가중치는 고정합니다. 발전소 간 창을 연결하거나 누락 시간을 임의의 값으로 채우지 않습니다.

| 구간 | 사용 목적 |
|---|---|
| Train | 모델 학습, CNN 결측 대체값 계산 |
| Validation | 모델별 특징·입력창·Optuna 설정 선택, early stopping |
| Calibration 앞부분 | 확정된 두 모델의 하이브리드 비중 학습 |
| Calibration 뒷부분 | 별도 purge 뒤 단일 모델과 Hybrid의 채택 평가 |
| Test | 선택을 확정한 뒤 최종 결과 보고 |

분할 기준은 모든 발전소의 원본 시간축으로 고정됩니다. 기본 구간 비율은 60/15/10/15이고 경계에는 purge가 적용되므로 실제 행 비율은 달라집니다. 외부 구간 purge는 최소 예측 기간 이상이며, 하이브리드 학습/선택 사이에도 별도 purge를 둡니다. Calibration을 알림 임계값 조정에 재사용하지 않습니다. 기존 이상치 분석과 새 벤치마크의 화면·산출물은 구분됩니다.

## 모델별 탐색과 선택

`config/experiments/optimized.json`은 CLI가 실제로 읽는 실행 설정입니다.

- XGBoost: 관측 기상 포함/제외 특징 조합 각각에서 Optuna 최대 15회.
- CNN-BiLSTM: 위 특징 조합 × 24/168시간 입력창 각각에서 Optuna 최대 5회. 네트워크 크기·커널·dropout·학습률 등은 모델 설정의 탐색 공간을 사용합니다.
- 각 후보의 시간 예산은 3,600초이며 최종 학습 시간은 별도입니다. 전체 설정은 장시간 실행이 될 수 있습니다.
- 같은 기간의 후보들은 공통 발전소·목표·시작 시각의 Validation 행에서 선택합니다. 서로 다른 모델의 특징 수와 입력창을 강제로 일치시키지 않습니다.
- 기본 공통 행 충족률 95% 미만이면 오류로 종료합니다. 제외된 행 수와 충족률을 남기며 점수가 좋은 행만 골라 비교하지 않습니다.
- Hybrid는 별도 채택 구간의 MAE가 최선 단일 모델보다 엄격히 낮을 때만 선택합니다. 최소 상대 개선율은 설정할 수 있습니다. 동률은 단일 모델을 유지합니다.
- Test 결과가 달라도 선택을 되돌리지 않습니다. Test에서 개선이 유지됐는지와 기준선 대비 성능을 별도로 표시합니다.

이 결과는 **설정한 후보와 예산 안에서 검증 점수가 가장 좋은 모델**입니다. 전역 최적 성능이나 다른 연도·발전소에 대한 우위를 보장하지 않습니다. 모든 후보는 같은 데이터의 시간 분할을 이용합니다. 탐색 과정에서 생성된 후보 Test 산출물은 선택에 읽지 않으며, 화면에는 선택된 후보의 Test만 사용합니다.

## 이름과 산출물 규칙

| 위치 | 역할 |
|---|---|
| `evaluation/experiment_config.py` | 실험 스키마 검증, 후보별 실행 설정 구성 |
| `evaluation/forecast_samples.py` | 시작시각·목표시각 정렬과 연속 관측창 검증 |
| `jobs/benchmark_job.py` | 후보 실행, Validation 선택, 다음 단계 연결 |
| `evaluation/model_selection.py` | 하이브리드 학습·채택·최종 평가 |
| `reporting/benchmark_analytics.py` | 완료된 산출물 검증 후 화면 데이터 구성 |
| `dashboard/src/benchmark.js` | 기간별 비교와 선택 근거 표시 |
| `artifacts/benchmarks/<run>/experiment.json` | 실제 적용한 실험 설정 |
| `artifacts/benchmarks/<run>/candidates/` | 기간/모델/특징·입력창별 checkpoint·전처리·설정·예측 |
| `artifacts/benchmarks/<run>/horizon_<h>h/base_selection.json` | Test를 읽기 전에 확정한 모델별 선택 |
| `artifacts/benchmarks/<run>/horizon_<h>h/selection.json` | 채택 이유, 구간, 검증/Test 지표, 파일 SHA256 |

함수명은 동작과 목적을 나타내는 snake_case, 소스 파일명은 책임을 나타내는 이름을 사용합니다. 새 모듈의 목적은 `config/architecture/modules.json`에 등록하고 `tools/update_code_index.py`, `tools/build_dashboard_assets.py`, `tools/check_structure.py`로 색인·번들·규칙을 검증합니다. 가중치·CSV·실행 JSON은 `src` 안에 생성하지 않습니다.

화면은 기간별로 XGBoost/CNN/Hybrid/기준선의 MAE·RMSE·R², 채택 이유와 실제/예측 곡선을 표시합니다. 발전소 행을 모은 지표와 지역·대상 발전소 전체의 시간별 합산 지표를 구분합니다. 음수 R²를 0으로 바꾸지 않으며 정의할 수 없는 R²는 비워 둡니다. 누락·중복·정답 변경·파일 hash 불일치가 있는 결과와 smoke 실행은 성능 화면에 올리지 않습니다.

## 실데이터 실행 검증의 범위

CI의 `tools/run_observed_benchmark_pilot.py`는 저장소의 공식 원본으로 Gold를 생성하고, 관측 충족률과 연속 입력창 충족률 기준으로 발전소 1개 연도를 선정합니다. 실제 자료로 1/24시간, CNN 입력창 24/168시간을 실행하되 CPU 자원과 탐색 횟수를 제한합니다. `tools/verify_benchmark_model_artifacts.py`는 저장된 가중치·전처리값을 다시 불러와 재학습 없이 전체 Test 예측을 재현합니다. 화면에도 `bounded_observed_pilot`로 표시합니다. 이는 실데이터 연결·학습·평가의 제한된 실험이며 전체 발전소의 충분한 최적화 결과는 아닙니다.

기존 구조 변경의 합성 데이터 동등성 CI는 새 예측 설계의 정확도 근거가 아닙니다. 과거 `92573c04`에 보관된 CNN/XGBoost checkpoint와 원래 평가 자료의 성능 재현은 별도 과제이며 이번 결과와 동일하다고 주장하지 않습니다. `controlled.json` 등 과거 실험 문서는 실행 가능한 현재 실험 스키마가 아니므로 `benchmark`가 거절합니다.

ASOS 키는 프로젝트 루트 `.env.local`의 `KMA_ASOS_SERVICE_KEY`에 둡니다. endpoint는 `https://apis.data.go.kr/1360000/AsosHourlyInfoService/getWthrDataList`입니다. 저장된 과거 관측자료로 실험할 때는 API 키를 사용하지 않습니다.
