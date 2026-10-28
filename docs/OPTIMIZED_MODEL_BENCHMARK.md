# 실측 자료 기반 모델 최적화와 비교

목적은 취업 포트폴리오에서 공식 발전량·기상청 ASOS 관측자료를 처리하고, 모델별 최적화 과정과 하이브리드 채택 근거를 재현 가능한 화면으로 보여주는 것입니다. 예보 API 가입은 필요하지 않습니다. 기본 1·24·72시간은 비교 과제의 예시이며 24시간으로 고정하지 않습니다. CNN-BiLSTM이 장기 예측에 반드시 우수하다고 가정하지 않습니다.

2026-09-16 사용자 결정에 따라 **각 모델의 성능 개선을 우선**합니다. 과거 CNN과 동일한 가중치·구조·점수 재현은 신규 모델 개발의 선행 조건이 아닙니다. 모델별 특징, 과거 입력 길이, 전처리, 학습 횟수와 탐색 예산은 다르게 구성할 수 있습니다. 비교할 때는 같은 예측 기간의 발전소·시작시각·목표시각·정보 이용 시점과 평가 지표를 맞춥니다. XGBoost 1시간 점수와 CNN 72시간 점수를 직접 비교해 승자를 정하지 않습니다.

**168시간 과거 입력창과 168시간 뒤 예측은 서로 다른 설정**입니다. 긴 과거 패턴을 학습하는 능력은 먼 미래의 발전량 정확도를 보장하지 않습니다. 단기 CNN 열세나 장기 CNN 우세를 목표 결과로 정하지 않고, 기간별 검증 결과에 따라 판단합니다. 제한된 3 epoch 실험도 각 모델의 최적 성능 근거로 사용하지 않습니다.

## 실행

```powershell
python app.py prepare-data --no-collected-downloads
python app.py benchmark --plan
python app.py benchmark --config config/experiments/optimized.json
python app.py serve-dashboard
```

`benchmark`는 완료 후 대시보드 데이터를 갱신합니다. `--plan`은 실제 후보·피처·탐색 예산을 출력하며 학습하지 않습니다. `--smoke`는 1시간·24시간 입력창·소량 데이터로 연결만 점검하고 대시보드의 성능 비교에서 제외합니다.
날짜 고정 설정에 `--smoke`를 사용하면 작은 표본에 맞는 비율 분할로 바꾸며, 요청 날짜와
실행 분할 변경을 계획·산출물에 기록합니다. 원래 날짜 분할의 정확도 검증으로 해석하지 않습니다.

## 평가 과제

목표 발전량 시각이 `t`, 예측 기간이 `h`이면 예측 시작 시각은 `t−h`입니다. 각 모델은 `h`시간 뒤 한 시점의 시간당 발전량을 예측합니다. 화면의 시계열은 시작 시각을 이동하며 만든 예측들을 연결한 것입니다. XGBoost는 `t−h`의 관측 피처, CNN은 `t−h`까지 포함하는 연속 시간창을 사용합니다. 예측 CSV에는 목표시각, 시작시각, 기간, 발전소 ID, 실측값, 예측값과 시작시각 발전량을 유지한 기준선이 저장됩니다. 미래에 관측된 날씨를 과거 예측 입력으로 사용하지 않습니다.

이는 이미 확보한 과거 관측자료로 수행하는 rolling-origin 평가입니다. 시각 `t−h`의 시간 관측이 그 시각에 사용 가능하다고 가정하며, 실제 API 발행 지연은 재현하지 않습니다. Test 진행 중 새로 관측된 이전 시점 자료를 입력으로 갱신할 수 있지만 모델 가중치는 고정합니다. 발전소 간 창을 연결하거나 누락 시간을 임의의 값으로 채우지 않습니다.

| 구간 | 사용 목적 |
|---|---|
| Train | 모델 학습, CNN 결측 대체값과 입력 표준화 통계 계산 |
| Validation | 모델별 특징·입력창·Optuna 설정 선택, early stopping |
| Calibration 앞부분 | 확정된 두 모델의 하이브리드 비중 학습 |
| Calibration 뒷부분 | 별도 purge 뒤 단일 모델과 Hybrid의 채택 평가 |
| Test | 선택을 확정한 뒤 최종 결과 보고 |

분할 기준은 모든 발전소에 동일한 전역 calendar boundary로 고정합니다. 발전소별로 60/15/10/15 cutoff를 따로 계산하지 않습니다. 확보 시작시점 차이 때문에 발전소별 실제 행 비율은 달라질 수 있지만 Train/Validation/Calibration/Test의 시간 의미는 동일합니다. 공통 경계를 만족하지 못하는 short-history 발전소는 main cohort에서 제외하거나 별도 분석 대상으로 보고하며, 해당 발전소만 Test 구간을 이동하지 않습니다. 외부 구간 purge는 최소 예측 기간 이상이며, 하이브리드 학습/선택 사이에도 별도 purge를 둡니다. Calibration을 알림 임계값 조정에 재사용하지 않습니다. 기존 이상치 분석과 새 벤치마크의 화면·산출물은 구분됩니다.

### 날짜 고정 후보와 학습 전 점검

`config/experiments/optimized.json`은 이제 Train 종료를 2024-06-30 23시,
Validation 종료를 2024-09-30 23시, Calibration 종료를 2024-12-31 23시,
Test 종료를 2025-12-31 23시로 고정한 전역 calendar split을 기본 계약으로 사용합니다.
이 경계는 2026-09-17 coverage audit를 통과한 기존 날짜 고정 후보를 canonical development-evaluation
split으로 승격한 것입니다. 2024년 상반기까지 Train에 포함해 신규 발전소의 이력을 확보하고
2025년 확보분을 동일 Test 기간으로 평가합니다. Validation이 3분기에 치우친다는 계절 일반화
한계는 그대로 보고합니다. 현재 로컬 Gold가 확장되었으므로 본 학습 전 동일 경계를 재감사하며,
충분성 문제가 있으면 Test 성능을 보기 전에 하나의 공통 경계 세트만 새 버전으로 수정합니다.

```powershell
python tools/audit_temporal_split.py --config config/experiments/optimized.json --data file/standardized/model_ready_parts --output artifacts/audits/observed_calendar_split.json
python app.py benchmark --config config/experiments/optimized.json --plan
```

첫 명령은 학습 없이 분할별 실제 발전소·월·계절·Train 이력 없는 발전소와 purge 제외 행을
기록합니다. 최종 모델 표본 수는 예측 기간·연속 입력창·피처 조건 때문에 더 줄어들 수 있습니다.
네 종료 필드는 모두 지정해야 하며 시간대 없는 KST의 **포함되는 종료 시각**입니다.
날짜만 쓰면 해당 날짜 00시이므로 마지막 시간까지 포함하려면 `T23:00:00`을 명시합니다.
데이터가 추가돼도 경계는 이동하지 않고 Test 종료 이후 행은 이 실행의 평가에서 제외됩니다.

이미 확인한 pilot Test와 새 후보 구간은 개발 평가로 취급합니다. 날짜를 바꿨다는 이유만으로
새로운 미관측 확증 Test가 되지 않습니다. 정기 재학습이나 드리프트 자동화는 이번 범위에
포함하지 않습니다.

2026-09-17, 저장소 원본 `8d5e64e7a2c484cb2068a292e2eae5c122c3cf43`의 기존 표준 파일
88개와 ASOS를 보완된 코드로 결합한 점검에서는 Gold 613,340행, 태양광 595,820행,
태양광 학습 적격 566,540행이 유지됐습니다. Windows 로컬에만 있는 추가 수집물은 포함하지
않았습니다. 기존 Gold와 모든 발전소·시각 키 및 발전량 원값 613,340행의 정확한 일치를
확인했습니다. 대응 ASOS 관측소·시각이 없는 2,787행은 결측 사유를 기록하며 주야 판정도
미상으로 유지했습니다. 아래는 168시간 경계 간격 적용 후 **입력창 생성 전 관측 정답 행 수**입니다.

| 구간 | 실제 관측 기간 | 행 | 발전소 |
|---|---|---:|---:|
| Train | 2017-01-01 ~ 2024-06-30 | 413,708 | 18 |
| Validation | 2024-07-08 ~ 2024-09-30 | 32,640 | 16 |
| Calibration | 2024-10-08 ~ 2024-12-31 | 32,568 | 16 |
| Test | 2025-01-08 ~ 2025-12-31 | 79,560 | 16 |

경계 간격으로 제외된 행은 8,064개이며, 평가 대상 16개 발전소 모두 Train 관측 이력이
있었습니다. 이는 모든 발전소가 Train 전체 기간을 채웠다는 뜻은 아닙니다. 특히 2025년은
발전소별 수집 종료일이 달라 연중 관측이 있는 발전소는 일부뿐입니다. 2026년 발전량 2,832행은
대응 ASOS 연간 파일이 없어 계속 보류했습니다. 분할의 적절성과 모델 성능은 별개의 검증입니다.

이번 보완의 로컬 검증은 ASOS/API·관측 전처리·CNN 수치 변환·날짜 분할 관련 unittest
55개, 코드 구조 검사, Python 컴파일, dashboard JavaScript 구문 검사를 통과했습니다.
이 초기 점검 당시에는 Torch/Optuna/XGBoost가 없어 프레임워크 실행을 검증하지 않았습니다.
이후 독립 환경에서 수행한 실행 검증은 아래에 별도로 기록하며 과거 CI 기록과 구분합니다.

### 실제 모델 입력창 적용 후 표본

`tools/audit_forecast_readiness.py`는 같은 Gold에 실제 모델의 품질 필터·정확한 예측 시작시각·
CNN 연속 시간창 조건을 적용합니다. 미래 발전량과 누락 시간은 채우지 않습니다. 모델별로
다른 학습 표본을 유지하고, 비교할 Validation/Calibration/Test의 공통 시각을 확인합니다.

```powershell
python tools/audit_forecast_readiness.py --config config/experiments/optimized.json --data file/standardized/model_ready_parts --output artifacts/data_audits/forecast_readiness.json
```

앞의 동일 Gold에서 확인한 **Train 표본 수**는 다음과 같습니다. 기상 포함/제외 피처셋은
동일한 표본을 사용하며, 결측 입력 때문에 유리한 발전소·시각만 남기지 않습니다.

| 예측 기간 | XGBoost | CNN 입력 24시간 | CNN 입력 168시간 |
|---|---:|---:|---:|
| 1시간 뒤 | 413,686 | 413,180 | 410,012 |
| 24시간 뒤 | 413,224 | 412,718 | 409,550 |
| 72시간 뒤 | 412,312 | 411,806 | 408,638 |

세 기간 모두 Validation 32,640행과 Test 79,560행의 대상 시각이 후보 전체에서 일치합니다.
Calibration의 전체 후보 공통 비율은 약 99.49%로 설정한 95% 기준을 통과했습니다.
이 비율은 **비교 가능한 표본의 비율이며 예측 정확도가 아닙니다**. 실제 benchmark는
Validation에서 모든 후보를 비교한 뒤, 확정된 두 모델의 Calibration/Test만 비교하므로
감사 도구의 전체 후보 Calibration/Test 교집합은 더 보수적인 사전 점검입니다.

24시간 뒤 예측·168시간 입력 CNN의 실제 Train 입력에 쓰이는 관측행 413,228개에서
강수 90.667%, 일사 68.860%, 설치각 64.903%, 일조 45.034%가 결측이었습니다. 이 비율을
센서 고장률로 해석하지 않습니다. 야간·원천 비관측·메타데이터 누락 등이 함께 포함됩니다.
Train 전체가 결측이거나 관측값이 상수인 피처는 없었습니다. 결측 표시를 보존하고 모델별
기상 포함/제외 후보를 Validation에서 비교하는 기존 설정을 유지합니다.

## CI 없이 로컬에서 실행

본 학습은 **사용자 Windows 로컬 GPU에서만** 진행합니다. ChatGPT 작업 환경과 GitHub
Actions에서 본 학습을 시작하지 않습니다. **이 변경 코드와 학습 의존성이 설치된 로컬**에서
기존 Gold를 사용해 다음 한 명령으로 표본 점검 → 본 학습 → 저장 모델 재예측 → 화면 데이터
생성을 수행합니다. Windows에서도 동일한 Python 도구를 사용합니다.

```powershell
.\.venv\Scripts\python.exe tools/run_local_benchmark.py --data file/standardized/model_ready_parts
```

`--config` 기본값은 날짜 고정 후보 `observed_calendar_candidate.json`입니다. `--data`로
기존 Gold 위치를 지정하며 탐색 예산·후보·분할 날짜는 바꾸지 않습니다. 도구는 의존성과
Gold v2 manifest를 확인하고 실패 단계·실제 환경을 `artifacts/verification/local_benchmark/`
아래 실행별 `status.json`에 기록합니다. 패키지 설치나 Gold 재생성, CI 호출은 수행하지 않습니다.
CUDA GPU를 사용할 수 없으면 본 학습과 `--preflight-only` 모두 의존성 확인 단계에서
종료하며 CPU로 자동 전환하지 않습니다. CNN은 GPU를 사용하고 XGBoost의 현재 구현은
로컬 CPU를 사용합니다. 기본 CPU 연산 스레드는 4개이며 `--threads`로 지정합니다.
재학습 없는 `--replay-run`은 기존 CPU 검증을 허용합니다. GPU 학습과 CPU 재예측 간
수치 오차로 엄격한 일치 검사가 실패할 경우 그 결과를 성능 저하로 단정하지 않으며,
오차 허용치를 임의로 늘려 통과시키지 않습니다. 사용자 로컬 GPU 실행은 아직 검증하지 않았습니다.

```powershell
# 학습 전 의존성·Gold·표본 점검만 수행
.\.venv\Scripts\python.exe tools/run_local_benchmark.py --preflight-only
# 학습 완료 후 재예측 또는 화면 단계만 다시 수행
.\.venv\Scripts\python.exe tools/run_local_benchmark.py --replay-run artifacts/benchmarks/<완료된_실행>
```

재예측 대상은 학습 서비스가 실제 반환한 실행 폴더입니다. 재예측이 통과한 뒤 대시보드를
갱신하고 해당 실행이 표시되는지 확인합니다. `--preflight-only` 통과는 정확도 검증이 아닙니다.

2026-09-17 별도 Python 3.12.14 환경에 PyTorch 2.14.0+cu130, XGBoost 3.4.1,
Optuna 5.0.0을 설치했습니다. GPU는 없어 CPU로 실제 프레임워크 통합 검사 54개와
subtest 31개를 통과했습니다. 검사에는 학습·체크포인트 재개·저장 전처리·날짜 분할이
포함됩니다. 이 결과만으로 본 실데이터 학습 완료나 정확도 향상을 주장하지 않습니다.

단계별 명령이 필요한 경우 다음과 같이 실행합니다. 데이터 원본·API 키를 외부 서비스로
전송하는 단계는 없습니다.

```powershell
Set-Location C:\ESG_Project1
$pvPython = ".\.venv\Scripts\python.exe"
function Invoke-PvStep {
    & $pvPython @args
    if ($LASTEXITCODE -ne 0) { throw "명령이 실패했습니다. 원인 확인 후 해당 단계부터 다시 실행하세요." }
}

Invoke-PvStep -c "import torch, xgboost, optuna, pytest; print('학습 환경 확인 완료')"
Invoke-PvStep -m pytest -q
Invoke-PvStep tools/check_structure.py

# 기존 Gold가 이전 전처리로 생성됐다면 보완된 코드로 한 번 생성합니다.
Invoke-PvStep app.py prepare-data
Invoke-PvStep tools/audit_forecast_readiness.py --config config/experiments/optimized.json --output artifacts/data_audits/forecast_readiness.json
Invoke-PvStep app.py benchmark --config config/experiments/optimized.json --plan

# 본 학습: 후보 18개, Optuna 최대 150회, 최종 학습 18회입니다.
Invoke-PvStep app.py benchmark --config config/experiments/optimized.json
```

의존성 확인이나 감사가 실패하면 다음 단계로 진행하지 않습니다. 후보별 3,600초는 Optuna
탐색 예산이며 진행 중 trial·최종 학습·예측 생성 시간은 더 걸릴 수 있습니다. 이미 완료한
pilot과 반복 재학습 검사를 선행 조건으로 다시 실행할 필요는 없습니다.

학습이 중단되면 데이터·설정·체크포인트·Optuna DB를 보존하고 같은 benchmark 명령을
재실행합니다. 새 결과 폴더를 만들면서 기존 학습 상태를 재사용하도록 실행별 출력 경로를
학습 식별값에서 분리했습니다. 같은 외부 실행 manifest를 이어 쓰는 기능은 아닙니다.
데이터·피처·분할·예측 기간·seed·CNN 입력 변환 계약이 바뀌면 다른 학습으로 취급합니다.

완료 후 출력된 실행 폴더를 사용해 **재학습 없이** 저장 모델 재예측을 검증합니다.

```powershell
Invoke-PvStep tools/verify_benchmark_model_artifacts.py "artifacts/benchmarks/<완료된 실행 폴더>" file/standardized/model_ready_parts --output artifacts/verification/calendar_model_replay.json
```

CSV/GZIP 단일 파일과 분할 CSV/GZIP 폴더를 같은 로더 계약으로 처리합니다. 분할 데이터
식별값은 수정 시각 대신 상대 경로·실제 파일 내용 SHA256과 Gold manifest를 사용합니다.
파일 복사나 시간 정보 변경만으로 식별값이 달라지지 않으며, 같은 크기의 내용 변경도
탐지합니다. 과거 크기·수정 시각 기반 분할 식별값 및 실행 폴더 의존 벤치마크 체크포인트는
새 계약에 자동 호환시키지 않습니다. 단일 파일 SHA256 및 일반 `train` 출력 경로 계약은
유지합니다. 이번 수정의 식별·로딩 단위 검증과 실제 프레임워크 재개 실행 검증은 구분합니다.

초기 로컬 실행 경로 보완 후 영향 범위 unittest 56개, 구조 검사, Python 컴파일, 날짜 고정
실험의 `--plan` 실행이 통과했습니다. 실제 Gold 18개 후보의 표본 감사도 통과했습니다.
이 초기 단계에서는 GitHub Actions와 실제 프레임워크 학습·재예측을 실행하지 않았습니다.
이후 의존성 설치와 실제 본 학습의 진행 기록은 [실행 기록](OBSERVED_BENCHMARK_RESULT.md)에
구분해 남깁니다.

## 모델별 탐색과 선택

`config/experiments/optimized.json`은 CLI가 실제로 읽는 실행 설정입니다.

- XGBoost: 관측 기상 포함/제외 특징 조합 각각에서 Optuna 최대 15회.
- CNN-BiLSTM: 위 특징 조합 × 24/168시간 입력창 각각에서 Optuna 최대 5회. 네트워크 크기·커널·dropout·학습률·시퀀스 요약 방식(`readout`)은 모델 설정의 탐색 공간을 사용합니다.
- 각 후보의 시간 예산은 3,600초이며 최종 학습 시간은 별도입니다. 전체 설정은 장시간 실행이 될 수 있습니다.
- 같은 기간의 후보들은 공통 발전소·목표·시작 시각의 Validation 행에서 선택합니다. 서로 다른 모델의 특징 수와 입력창을 강제로 일치시키지 않습니다.
- 기본 공통 행 충족률 95% 미만이면 오류로 종료합니다. 제외된 행 수와 충족률을 남기며 점수가 좋은 행만 골라 비교하지 않습니다.
- Hybrid는 별도 채택 구간의 MAE가 최선 단일 모델보다 엄격히 낮을 때만 선택합니다. 최소 상대 개선율은 설정할 수 있습니다. 동률은 단일 모델을 유지합니다.
- Test 결과가 달라도 선택을 되돌리지 않습니다. Test에서 개선이 유지됐는지와 기준선 대비 성능을 별도로 표시합니다.

이 결과는 **설정한 후보와 예산 안에서 검증 점수가 가장 좋은 모델**입니다. 전역 최적 성능이나 다른 연도·발전소에 대한 우위를 보장하지 않습니다. 모든 후보는 같은 데이터의 시간 분할을 이용합니다. 탐색 과정에서 생성된 후보 Test 산출물은 선택에 읽지 않으며, 화면에는 선택된 후보의 Test만 사용합니다.

## CNN 개선의 근거와 검증 순서

현재 코드의 `last_output`은 BiLSTM 출력의 마지막 시점만 사용합니다. 이는 정방향 최종 상태와 역방향의 마지막 입력 시점 상태를 연결하며, 역방향으로 전체 창을 읽은 최종 상태와는 다릅니다. 신규 `final_hidden`은 마지막 LSTM 층의 정·역방향 최종 상태를 함께 사용합니다. 이 구분은 [PyTorch LSTM 공식 문서](https://docs.pytorch.org/docs/2.14/generated/torch.nn.LSTM.html)의 `h_n` 설명에 근거합니다. 예측 시작시각 이전의 입력창 안에서만 양방향 연산을 수행합니다.

기본 최적화 설정은 두 요약 방식을 Validation에서 탐색합니다. 탐색 공간을 생략한 신규 학습과 Optuna를 사용하지 않는 학습은 `final_hidden`을 사용합니다. 저장 모델에는 사용한 `readout`을 기록하고, 이 필드가 없는 과거 모델은 `last_output`으로 복원합니다. 가중치 크기가 같더라도 계산 의미가 달라지므로 과거 가중치를 새 요약 방식으로 임의 해석하지 않습니다. 기존 Optuna 탐색 결과와도 학습 계약을 구분합니다.

현재 구현은 구조상 근거가 있는 개선 후보이며 **정확도 향상 자체를 입증한 것은 아닙니다**.
CNN은 실제 Train 입력창에 사용된 중복 없는 행에서 중앙값 대체 후 평균·표준편차를 적합합니다.
결측 마스크는 이진 값으로 유지하고 타깃 MWh는 변환하지 않습니다. 통계·피처 순서·변환 버전을
모델과 함께 저장하며 모든 checkpoint 평가 경로는 저장된 통계와 시간 경계를 재사용합니다.
과거 median-only 모델은 표준화를 적용하지 않는 명시적 호환 경로로 읽습니다. 저장 전처리가
없는 모델의 평가 데이터를 이용해 통계를 재적합하지 않습니다. 입력창 길이와 충분한 학습,
early stopping은 이후 성능 실험에서 검증합니다. XGBoost는 native NaN과 별도 트리 탐색을
유지하며 Hybrid는 독립 채택 구간에서 최선 단일 모델보다 개선된 경우에만 선택합니다.

Gold의 `solar-observed-preprocessing.v2`는 발전량 원값을 보존하면서 무효값의 lag/rolling
재유입을 막고 ASOS 변수별 관측 여부·QC·결측 사유를 기록합니다. 관측소·시각 자체가 없는
경우는 `station_hour_missing`으로 구분합니다. 결측 강수·야간 일사를 임의로 0으로 채우지
않으며, 기상 이력은 당시 유효한 관측소 좌표를 사용합니다. 원본 파일은 수정하지 않습니다.

이미 점검한 pilot Test는 개발 과정의 진단 기록입니다. 이를 다시 실행한 점수를 새로운 확증 평가로 표시하지 않습니다. 본 실험의 최종 성능 주장은 아직 열어보지 않은 시간 구간을 확보하거나, 시간순 다중 구간 검증과 별도 최종 Test를 사전에 고정한 뒤 평가합니다.

## 이름과 산출물 규칙

| 위치 | 역할 |
|---|---|
| `evaluation/experiment_config.py` | 실험 스키마 검증, 후보별 실행 설정 구성 |
| `evaluation/forecast_samples.py` | 시작시각·목표시각 정렬과 연속 관측창 검증 |
| `evaluation/split_audit.py` | 학습 없는 시간 분할·계절·발전소 커버리지 감사 |
| `evaluation/forecast_readiness.py` | 모델별 실제 입력창 표본·공통 평가 시각·Train 결측 점검 |
| `models/cnn_bilstm/input_preprocessing.py` | Train 통계 적합과 저장 계약에 따른 입력 변환 |
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

완료한 실행의 입력·탐색 예산·점수·저장 모델 재검증은 [공식 실측 자료 실행 기록](OBSERVED_BENCHMARK_RESULT.md)에서 확인합니다. 설정에 정의한 전체 탐색과 실제 완료한 제한 실험을 구분합니다.

CI의 `tools/run_observed_benchmark_pilot.py`는 저장소의 공식 원본으로 Gold를 생성하고, 관측 충족률과 연속 입력창 충족률 기준으로 발전소 1개 연도를 선정합니다. 실제 자료로 1/24시간, CNN 입력창 24/168시간을 실행하되 CPU 자원과 탐색 횟수를 제한합니다. `tools/verify_benchmark_model_artifacts.py`는 저장된 가중치·전처리값을 다시 불러와 재학습 없이 전체 Test 예측을 재현합니다. 화면에도 `bounded_observed_pilot`로 표시합니다. 이는 실데이터 연결·학습·평가의 제한된 실험이며 전체 발전소의 충분한 최적화 결과는 아닙니다.

저장 모델 재예측과 새로 학습한 모델의 재현성은 별개입니다. CI는 `config/environments/benchmark_cpu_py311.constraints.txt`로 CPU 패키지 버전을 고정하고 실제 CPU/Torch 연산 설정을 남깁니다. 이어 `tools/verify_benchmark_retraining.py`가 같은 실측 입력으로 새 체크포인트·Optuna DB에서 학습을 반복해 모든 후보의 예측과 선택을 대조합니다. 이미 본 Test의 반복은 재현성 검사이며 새로운 확증 평가가 아닙니다. 실행 환경을 바꾼 CNN 결과가 달랐던 기록과 한계도 [실행 기록](OBSERVED_BENCHMARK_RESULT.md)에 보존합니다.

기존 구조 변경의 합성 데이터 동등성 CI는 새 예측 설계의 정확도 근거가 아닙니다. 과거 `92573c04`에 보관된 CNN/XGBoost checkpoint와 원래 평가 자료의 성능 재현은 별도 과제이며 이번 결과와 동일하다고 주장하지 않습니다. `controlled.json` 등 과거 실험 문서는 실행 가능한 현재 실험 스키마가 아니므로 `benchmark`가 거절합니다.

ASOS 키는 프로젝트 루트 `.env.local`의 `KMA_ASOS_SERVICE_KEY`에 둡니다. endpoint는 `https://apis.data.go.kr/1360000/AsosHourlyInfoService/getWthrDataList`입니다. 저장된 과거 관측자료로 실험할 때는 API 키를 사용하지 않습니다.
