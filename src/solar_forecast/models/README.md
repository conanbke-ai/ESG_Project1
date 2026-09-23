# 모델 구현 위치

| 폴더 | 모델 | 먼저 볼 파일 |
| --- | --- | --- |
| `cnn_bilstm/` | CNN-BiLSTM 시계열 신경망 | `network.py` 구조 → `sequence_data.py` 입력 → `trainer.py` 실행 연결 |
| `xgboost/` | XGBoost 회귀 모델 | `trainer.py` 학습 → `optimization.py` 탐색 → `checkpoint.py` 재개 |
| `hybrid/` | Validation 기반 CNN/XGBoost 결합 | `dynamic_gate.py` 결합 규칙 → `experiment.py` 평가 |
| `shared/` | 공통 지원 코드 | `checkpoint_store.py`, `optuna_study.py` |

이곳에는 Python 구현만 둡니다. 학습 조건은 프로젝트 루트의 `config/models/<model_id>.json`, 결과는 `artifacts/models/<model_id>/<run_id>/`, 재개 상태는 `artifacts/checkpoints/`에 저장합니다. `config/experiments/model_upgrade_candidates.json`의 후보 모델은 구현된 모델 목록과 구분합니다.

세부 파일·함수는 [코드 색인](../../../docs/CODE_INDEX.md), 생성 규칙은 [프로젝트 구조](../../../docs/PROJECT_STRUCTURE.md)를 참조하세요.
