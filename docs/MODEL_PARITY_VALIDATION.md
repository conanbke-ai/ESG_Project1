# 모델 구조 정비 전후 검증

목표는 코드 이름·배치 변경이 같은 입력의 학습 결과를 바꾸었는지 직접 확인하는 것이다. 전체 회귀 테스트 성공, 수치 동등성, 과거 실데이터 성능 재현은 각각 별개의 결과로 기록한다.

## 자동 실행하는 검증

`.github/workflows/structure.yml`의 `model-regression` job은 같은 Python 3.11 CPU 환경에서 전체 pytest와 변경 전후 비교를 실행한다. 변경 전 기준은 KOSPO 교정 직후이자 구조 정비 직전인 commit `477b01e9e4944330b7b17695b31d40c800819bf3`이다.

`tools/verify_model_reorganization.py`는 두 checkout을 별도 프로세스에서 읽는다. Python·NumPy·PyTorch 난수와 CPU 스레드를 고정하고, Optuna·checkpoint 재개를 끈 작은 학습 조건을 양쪽에 동일하게 적용한다. 검증용 데이터는 한 번 생성한 같은 CSV byte를 사용한다. 기본 실행의 데이터는 합성이며 실제 발전소·ASOS 관측 실적이 아니다. 기존 Gold, 기본 모델 설정, 보관 모델 파일을 변경하지 않는다.

검증 산출물은 CI의 `model-reorganization-evidence`에 저장한다. 입력·설정 hash, 코드 revision, 실행 환경, 분할별 예측과 비교 보고서를 함께 보관한다. 수치 오차 허용값은 절대·상대 각각 `1e-6`이며 실제 최대 차이도 보고한다.

```powershell
python tools/verify_model_reorganization.py --baseline-root C:\ESG_Project1_before --candidate-root C:\ESG_Project1 --output-dir artifacts/verification/model_parity_run1
```

두 폴더에는 각각 해당 코드가 있어야 하고 같은 인터프리터에 프로젝트 의존성이 설치되어 있어야 한다. 결과 폴더는 이전 실행 결과와 겹치지 않는 새 이름을 사용한다.

## 기존 예측 결과를 직접 비교

이미 보관한 변경 전후 예측 CSV가 있다면 다시 학습하지 않고 비교할 수 있다.

```powershell
python app.py compare-predictions C:\ESG_Project1_before\artifacts\models\cnn_bilstm\baseline_run\test_predictions.csv artifacts/models/cnn_bilstm/candidate_run/test_predictions.csv --report artifacts/verification/cnn_bilstm_prediction_parity.json
```

필수 열은 `plant_id`, `timestamp`, `y_true`, `y_pred`다. 선택 열 `region`, `split`이 한쪽에 있으면 다른 쪽에도 있어야 하며 내용이 일치해야 한다. 중복 키, 다른 평가 행 집합, 다른 정답, 비정상 시각, NaN·무한대는 실패다. 서로 다른 평가 구간의 점수만 보고 성능이 같다고 판단하지 않는다.

| 보고 값 | 해석 |
| --- | --- |
| `overall`, `per_plant`, `per_region` | 전체 및 그룹별 관측 행의 MAE·RMSE·R²와 candidate−baseline 차이 |
| `per_region` | 지역 안의 plant-hour 행을 묶은 지표. 시간별 지역 발전량 합계의 지표는 아니다. |
| `prediction_difference` | 예측값 절대 차이의 최대·평균과 허용 오차 초과 행 수 |
| `inputs` | 실제 읽은 파일 SHA256·행 수·시각 timezone 방식 |
| `production_accuracy_verified=false` | 이 비교만으로 운영 정확도나 과거 보고 성능을 인증하지 않음 |

상수 정답이나 표본 1개에서 정의되지 않는 R²는 `null`이다. 검토한 일광·계절·기상 분류 계약이 없으면 해당 조건별 분석을 제공한 것으로 표시하지 않는다. 비교 성공은 종료 코드 0, 수치 차이는 1, 잘못된 입력도 실패 종료한다. `--report`는 입력 CSV를 덮어쓸 수 없다.

## 과거 성능 재현에 필요한 기준

현재 저장소의 CNN은 Conv1d → BiLSTM → 단일 출력이고 원시 `generation_mwh`를 학습한다. 과거 메모에 있는 다른 층 크기, 24개 동시 출력, log1p 변환 모델과 같다고 가정하지 않는다. 또한 현재 고정 학습 경로는 설정의 seed를 직접 적용하지 않으므로 이 비교 실행기가 프로세스 시작 때 난수를 고정한다.

`horizon_hours=24`라는 manifest 필드만으로 24시간 동시 예측이 검증되었다고 판단하지 않는다. 현재 시퀀스는 직전 관측 구간 다음의 단일 정답을 예측한다. 실제 예보 시점에서 사용할 수 없는 미래 관측값을 사용하지 않았는지는 운영 평가에서 별도로 확인해야 한다.

과거 성능을 재현할 때는 백업의 모델 checkpoint, 전처리·target 변환 상태, 특징 순서, 원래 모델 설정, 고정 Gold와 manifest, 동일 테스트 구간과 예측 CSV, 당시 실행 환경을 기준으로 삼는다. ASOS를 추가해 데이터 구성이 바뀐 결과는 기존 데이터에서의 코드 동등성 결과와 따로 기록한다. 모델과 Gold가 현재 작업 환경에 없으면 그 성능은 미검증 상태로 남긴다.
