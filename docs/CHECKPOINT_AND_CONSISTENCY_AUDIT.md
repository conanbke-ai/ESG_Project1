# 학습 체크포인트 계약과 코드 정합성 검수

## 구현 범위

| 경로 | 저장 상태 | 기본 저장 단위 | 재개 위치 |
|---|---|---:|---|
| CNN Optuna trial | 모델, optimizer, epoch, Validation MAE, early-stopping wait, RNG | 1 epoch | 완료한 다음 epoch |
| CNN 최종 학습 | 위 상태와 best model state | 1 epoch | 완료한 다음 epoch |
| CNN adaptive 미세조정 경로 | 기반 모델, optimizer, bandit Q값/이력, best state, RNG | 1 epoch | 완료한 다음 epoch |
| XGBoost Optuna trial | Booster, 완료 round, early-stopping 최적 점수/대기 횟수 | 50 rounds | 저장된 Booster 이후 round |
| XGBoost 최종 학습 | Booster, 완료 round, early-stopping 최적 점수/대기 횟수 | 50 rounds | 저장된 Booster 이후 round |
| Optuna orchestration | study/trial/파라미터/heartbeat | trial 상태 변경 | 남은 trial 또는 실패 trial 1회 재시도 |

설정은 `config/models/*.json`의 `checkpoint`에서 관리합니다. `enabled=false`는 읽기와 저장을 모두
끄고, `resume=false`는 기존 상태를 읽지 않되 이번 실행의 상태는 계속 저장합니다.
레거시 `pipeline` 진입점도 실행별 산출물 폴더가 아니라 `PipelineConfig.output_dir/.checkpoints`를
사용하므로 새 실행 디렉터리가 만들어져도 같은 데이터·설정의 학습 상태를 찾을 수 있습니다.
이 경로의 Optuna study 역시 `PipelineConfig.output_dir/optimization.db`에 영속화합니다.

## 호환성과 원자성

체크포인트 디렉터리와 Optuna study는 모델명과 SHA-256 fingerprint로 격리됩니다. fingerprint에는
모델 profile, 피처·분할·모델 파라미터, `model_ready_manifest.json` hash와 파티션 inventory가
들어갑니다. 실행 시간 제한, 최대 누적 trial 수처럼 결과 의미를 바꾸지 않는 운영 예산은 제외합니다.
각 stage는 다시 모델 구조·선택된
파라미터·학습 길이·분할 정보를 signature로 검증합니다. 어느 값이든 다르면 명시적으로 실패하고
서로 다른 학습 상태를 이어 붙이지 않습니다.

PyTorch와 XGBoost는 임시 모델 파일을 완전히 쓴 뒤 같은 파일시스템에서 atomic replace합니다.
진행 metadata도 같은 방식으로 씁니다. XGBoost는 모델 자체의 실제 boosting round를 다시 읽으므로
모델 저장 직후 metadata 갱신 전에 중단되어도 중복 round를 학습하지 않습니다.

## 호출 흐름 검수 기준

```text
CLI train
  -> TrainingService (lock, run manifest)
    -> CnnBiLstmTrainer / XGBoostTrainer
      -> DatasetRepository (필터, dtype, memory limit)
      -> TemporalSplitter / CNN split loader
      -> OptunaStudyService (Validation-only selection)
      -> TrainingCheckpointStore (trial/final state)
      -> atomic model artifact
      -> TrainingService success manifest
```

- Train/Validation/Calibration/Test 역할은 기존 전역 시간 경계와 purge 계약을 유지합니다.
- Optuna 목적함수와 pruning에는 Validation만 들어가고 Test는 최종 평가에만 사용합니다.
- smoke 실행은 Optuna를 건너뛰지만 최종 학습 체크포인트 배선을 그대로 통과합니다.
- 완료 체크포인트는 최종 산출물과 성공 manifest 사이의 장애 창을 복구하기 위해 유지합니다.
- Optuna COMPLETE/PRUNED trial의 임시 모델 상태는 제거하고, FAILED trial 상태만 제한된 재시도를
  위해 보존합니다.

## 운영 한계와 관리

- 저장 주기 전에 중단되면 마지막 CNN epoch 또는 최대 50개 XGBoost round는 다시 계산합니다.
- SQLite heartbeat/stale callback은 현재 Optuna에서 experimental 경고를 내지만 공식 API를 사용합니다.
- 완료 체크포인트는 의도적으로 유지되므로 장기 운영 시 fingerprint별 보존 기간 또는 모델 registry
  승격 후 정리 정책을 추가해야 합니다.
- 분산 동시 쓰기를 위한 원격 lock/object-store transaction은 아직 범위 밖입니다. 현재는 단일 노드
  전역 학습 lock을 전제로 합니다.

## 재현 검수 명령

```bash
python -m pytest -q
python -m compileall -q src tests
python app.py --help
python app.py train xgboost --smoke
python app.py train cnn_bilstm --smoke
```

같은 smoke 명령을 한 번 더 실행하면 성공 manifest의 `checkpoint.resumed=true`와 동일한 완료
round/epoch를 확인할 수 있습니다. 데이터 또는 설정을 변경하면 fingerprint가 달라져 새 상태에서
시작해야 정상입니다.
