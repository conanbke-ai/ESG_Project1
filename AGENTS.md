# Development rules for AI/code agents

This repository uses the common TORI development-orchestration baseline from `conanbke-ai/Tori_Common_Project@943b567e7740d72727c45ba4c03d267e99dc4596`.

## Common entrypoint
- Read latest `main`, this file, relevant model/data docs, open PRs/branches, and actual code/config before changing behavior.
- Common routing: `policies/TORI_AI_DEVELOPMENT_ROUTER.md`.
- Common credit guard: `policies/TORI_AGENT_CREDIT_GUARD.md`.
- Common orchestration: `policies/TORI_AGENT_ORCHESTRATION_STANDARD.md`.
- Specialist definitions are logical ChatGPT roles under common `policies/agent_roles/`; no Cursor/plugin installation is required.

## Runtime routing
- Chat/connector first for dataset admission, temporal split, feature/model/experiment design, result analysis, docs, and repo inspection.
- Codex only for repository-local source changes plus targeted build/test/evaluation code.
- Long model training, Optuna and GPU-heavy benchmarking default to LOCAL_RUNTIME.
- Work is not used unless an actual dashboard/browser interaction acceptance task requires it.
- Do not duplicate one objective across execution surfaces.

## Data/ML contracts
- Preserve source provenance, raw bytes/hashes, registry/admission decisions, plant identity, weather mapping, and generation-weather overlap evidence.
- Dataset eligibility is decided before temporal split; do not impose a fixed start year without a canonical data-policy reason.
- Keep Train/Validation/Calibration/Test roles separate. Test is not a tuning surface.
- Use Data Guardian for lineage/admission/split risks and ML Reviewer for statistical validity/model comparison.
- Use Solar Forecast Reality Reviewer whenever operational forecast validity, issuance time, observed-vs-forecast weather, lag availability, latency, rolling updates, or deployability is claimed.
- Distinguish historical hindcast/evaluation evidence from true day-ahead operational evidence.
- Do not infer equipment failure/root cause from public generation/weather residuals without independent evidence.

## Validation
- Prefer targeted tests and smoke runs before broader regression.
- Record effective rows/sequences and split losses when eligibility/purge/window changes are material.
- Material experiment/model decisions must update canonical docs and portfolio history without overstating metrics or deployment readiness.
- Role triggers/regression baseline: common `docs/TORI_AGENT_TRIGGER_MATRIX.md` and `docs/TORI_ORCHESTRATION_REGRESSION_SCENARIOS.md`.
- Runtime recovery/retry/reconnect/checkpoint changes use Runtime Reliability review; prompt/model/provider/context changes affecting generative output use AI Output Evaluation review.
- Common routing regression result: `ORCHESTRATION_REGRESSION_PASS` (15/15 representative scenarios; first-pass routing gaps fixed in common policy).
- Mandatory user-facing gate: Product Flow QA. Any change to steps/forms/defaults/navigation/consent/validation/feedback/recovery/task-completion must be reviewed before implementation and before completion; it may require reordering/removing/defering steps for user convenience and error prevention.
- Professional QA baseline: common `policies/TORI_PROFESSIONAL_QA_STANDARD.md`.
- Every non-trivial feature/cross-module change starts with QA Lead/Test Architect PRECHECK: risk, acceptance traceability, mandatory QA roles, test cases/states, regression scope, and exit evidence.
- Behavior/state/validation changes require Functional & Regression QA; user-facing flows require Product Flow QA before implementation and before completion; Accessibility/Compatibility/Visual QA are selected by risk, not skipped because code builds.

## Tori Development Framework v2.0

- 신규 작업 및 다음 관련 checkpoint에서 공통 canonical `conanbke-ai/Tori_Common_Project`의 `policies/DEVELOPMENT_TEAM_OPERATING_POLICY.md` §26을 확인하고 적용한다. 기존 baseline은 역사적 기록으로 유지하며 실제 읽은 최신 SHA를 인계에 기록한다. 기존 제품 계약과 QA gate를 보존한다.

## Parallel Architecture Migration Policy

- 공통 운영 정책 §29를 적용한다. 기존 기능 브랜치 전체 완료를 선행 조건으로 두지 않으며, 독립적인 아키텍처 개선과 기능 개발을 병행한다. 기존 동등 workstream 우선 재사용, 필요한 경우 제품당 architecture 전용 branch 최대 1개 예외. 이전 branch 5개 제한 문구는 이 예외에 한해서 §29가 우선하며 WIP3/validation1/single writer/QA/보안/사용량 한도는 유지한다.
- 목표 구조/명명은 공통 §28과 `docs/TORI_Architecture_Blueprint.md`, `docs/TORI_Migration_Map.csv`를 참조하되 실제 제품 계약과 호출 관계를 확인한 뒤 변경한다. 신규 위반 방지와 기존 위반 점진 개선을 구분하고, 검증된 독립 변경만 통합한다. 실제 작성자/로컬 작업이 확인되지 않은 충돌 범위는 인계 또는 사용자 확인 후 진행한다.


## Configuration Management & Worktree Traceability

공통 운영 정책 §30을 다음 안전한 checkpoint부터 적용한다. 실제 checkout/worktree 또는 cloud workspace, 담당자/lease, base/HEAD/upstream, scoped diff, tested SHA, 원격 게시·main 통합·로컬 동기화 상태를 기존 ACTIVE_WORK/PR에 기록한다. 로컬 origin/*은 캐시이므로 실제 원격 확인과 구분한다. 기존 브랜치·미커밋 변경·원본 데이터·모델·산출물을 보호하며 무단 cleanup/force/reset/prune을 금지한다. 변경분 미확인 및 권한 차단은 Credit Guard §13에 따라 최소 사용자 조치로 인계하고 다른 독립 작업은 계속한다. 자동 검증 구현과 정책 채택은 별개다.


## Delegated Execution & Approval Boundaries

공통 운영 정책 §31의 A/B/C/D 위임 권한을 다음 안전한 checkpoint부터 적용한다. A 일반 개발·검증은 자율 실행, B 검증된 commit/non-force push/PR 및 엄격한 gate를 충족한 저위험 통합·안전 동기화는 조건부 자율 실행, C 비호환 계약·사업/보안 정책·고위험·파괴적 변경·비용 증가·운영 배포는 사용자 명시 승인, D 실제 접근/실행 차단은 최소 사용자 조치 인계다. 기존 작성자/lease, 필수 QA·보안·리뷰, 보호 브랜치 및 개별 PR 승인 범위를 유지한다. 일반적인 commit/push 승인을 반복 요청하지 않으며 실행 권한이 없는 작업을 수행했다고 주장하지 않는다.
