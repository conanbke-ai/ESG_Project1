# ACTIVE_WORK

## ESG Data Admission & Model Reuse Phase 2 — 2026-11-25

- 범위는 2024/2026 원천 완전성, 버전 지정 Gold 편입, 저장 모델 재사용, 독립 평가 준비다. 작업 트리는 `/workspace/ESG_Phase2`, 브랜치는 `feat/esg-data-admission-reuse-phase2`, 담당 writer는 이 Codex 채팅이다.
- 확인한 원격 main은 `3dbc598141e1c8fbae4290093aabb67d7ccddf70`이다. 최초 구현 기준은 기존 구조 개선 PR #2의 `44714db22e79ead6bdadaf8ad2522dfa7095a4c0`이다. 해당 브랜치의 Gold 감사, 학습 식별자, 분할 구현을 사용한다.
- 공통 정책 `conanbke-ai/Tori_Common_Project@f34e980c269ffdb121ad49a0d22e9c27ce45ad0b` §26·29–31을 확인했다. 기준 작업 트리는 깨끗했고 원본 checkout의 기존 브랜치와 저장 벤치마크를 참조한다. 현 작업은 main에 통합되지 않았다.
- 검증 결과와 재현 명령은 [Phase 2 결과](docs/ESG_DATA_ADMISSION_AND_MODEL_REUSE_PHASE2.md)에 있다. 다음 데이터 실행은 기존 Gold와 2026 원본·인덱스의 실제 파일 경로가 확인되는 시점에 해당 범위를 감사하는 것이다.

## POLICY_ACK — ChatGPT logical orchestration — 2026-09-18

- Common baseline: `conanbke-ai/Tori_Common_Project@943b567e7740d72727c45ba4c03d267e99dc4596`
- tags: `AGENT_ORCHESTRATION,AI_SURFACE,CREDIT,DATA,ML,SOLAR_FORECAST_REALITY,DOCUMENTATION`
- result: `ADOPTED`
- Runtime model: ChatGPT logical roles; no Cursor/plugin installation required.
- Surface rule: Chat/connector first; Codex only for repository-local implementation/build/test; long training/Optuna on LOCAL_RUNTIME; Work only if an actual dashboard/browser interaction acceptance task requires it.
- Product source-of-truth remains this repository's current main code/config/docs and data manifests.

## Current canonical workstream

- Dataset admission/registry/Gold eligibility and generation-weather continuity
- Temporal split integrity and effective per-plant subset sufficiency
- XGBoost / CNN-BiLSTM / Hybrid experiment design and evaluation
- Operational forecast-reality validation
- Local GPU training and Optuna
- Documentation / experiment / portfolio consistency

## Guardrails

- No fixed start-year rule unless current data policy requires it.
- No Test-set tuning.
- No production day-ahead claim from unavailable future observations.
- No automatic equipment-failure attribution from public residuals alone.
- No destructive data/branch/deployment action without explicit approval.

### ORCHESTRATION_V3_ACK — 2026-09-18
- Common baseline: `conanbke-ai/Tori_Common_Project@943b567e7740d72727c45ba4c03d267e99dc4596`
- Added logical gates: Runtime Reliability / AI Output Evaluation.
- Trigger/regression policy: ADOPTED.
- Product runtime code unchanged by this ACK.

### ROUTING_REGRESSION_ACK — 2026-09-18
- Common baseline: `conanbke-ai/Tori_Common_Project@943b567e7740d72727c45ba4c03d267e99dc4596`
- Result: `ORCHESTRATION_REGRESSION_PASS` — 15/15 representative scenarios.
- No product runtime code changed by this ACK.

### PRODUCT_FLOW_QA_ACK — 2026-09-23
- Common baseline: `conanbke-ai/Tori_Common_Project@943b567e7740d72727c45ba4c03d267e99dc4596`
- Result: `ADOPTED`
- User-facing changes require PRE_IMPLEMENTATION_FLOW_QA and PRE_COMPLETION_FLOW_QA; implementation may be reordered/trimmed when the current journey is unnecessarily difficult or risky.
- Product runtime code unchanged by this ACK.

### PROFESSIONAL_QA_V5_ACK — 2026-09-23
- Common baseline: `conanbke-ai/Tori_Common_Project@943b567e7740d72727c45ba4c03d267e99dc4596`
- Result: `ADOPTED`
- Non-trivial feature lifecycle: QA PRECHECK -> implementation -> Functional/Regression -> selected specialist QA -> Accessibility/Compatibility/Visual as applicable -> regression -> QA exit -> Acceptance.
- User-facing flows keep mandatory Product Flow QA before implementation and completion.
- Product runtime code unchanged by this ACK.
