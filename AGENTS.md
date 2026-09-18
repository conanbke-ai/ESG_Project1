# Development rules for AI/code agents

This repository uses the common TORI development-orchestration baseline from `conanbke-ai/Tori_Common_Project@98946df8f04685f0b2c9814363ec9fea8be25226`.

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
