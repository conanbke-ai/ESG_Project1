# ACTIVE_WORK

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
