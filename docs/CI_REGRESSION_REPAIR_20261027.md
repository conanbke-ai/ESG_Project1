# PR #2 CI regression repair — 2026-10-27

Baseline: `e908bb36ff7c60a9eaddb819795de5ae6f231d90`, Actions
[#40437298992](https://github.com/conanbke-ai/ESG_Project1/actions/runs/40437298992).
Continuation of [handoff](https://github.com/conanbke-ai/ESG_Project1/pull/2#issuecomment-6432100496).
Existing branch/base: `refactor/project-structure-asos` → `fix/kospo-plant-identity`.
Common policy: `c65cec53eac1b4c269a624321bbb90db8f5ca11f`, bootstrap/router and
[issue #6](https://github.com/conanbke-ai/Tori_Common_Project/issues/6).
Single writer: Work/Codex session `25b3692561fb`; earlier unpublished work isolated.

## Contract diagnosis and changes

| Failure | Evidence | Repair / acceptance |
|---|---|---|
| Stale code index | Canonical generator differs from tracked index | Run `python tools/update_code_index.py` before fixes and after test additions; no manual index edits. |
| Float32 CSV identity | `float32(0.12345679)` has bit pattern 1039980266; `.7g` becomes 1039980267 | Replace synthetic rounding fixture with actual pandas float32/float64 CSV roundtrips. Exact float32 truth comparison stays unchanged. New negative test rejects one representable step in both truth and persistence. |
| Three retraining failures | `run_dir` is a directory, but `artifact_path` requires a file | Resolve its `manifest.json` through the existing containment/file gate in selected-fit and search-only readers. Preserve relocation, exact truth, prediction tolerance, all candidates/splits and frozen selection. |
| Readiness fixture | Fixture lacks mandatory `selection_protocol` | Supply validation-only/common-row/pooled-MAE/frozen-Test protocol. Production validator stays unchanged. |
| Eligibility fixture | 100 rows with 40/10/10% leaves Train=40, Validation=40 | Check smaller/equal/larger cases (35/40/45), assert actual split counts. Only strict `non_train > train` rejects; production gate stays unchanged. |

No model architecture, optimizer budget, source data, schema, public API, temporal
split configuration, inference behavior, or Test-selection policy changed.
The only executable production-tool change is the retraining artifact reader.
No new abstraction or file move; canonical `preprocessing/` ownership and existing
compatibility shims remain in place.

## Executed local evidence

Linux x86_64, Python **3.11.17** (CI uses **3.11.16**, not claimed identical).
CPU dependencies installed from `benchmark_cpu_py311.constraints.txt`:
Torch 2.14.0+cpu, NumPy 2.4.6, pandas 3.0.5, scikit-learn 1.9.1,
XGBoost 3.2.0, Optuna 5.0.0, pytest 9.1.1. `uv pip check`: PASS.
Tests use `CUDA_VISIBLE_DEVICES=''`, OMP/MKL/OpenBLAS threads=1.

- Before source/fixture repair, four affected files: **6 failed, 28 passed,
  13 subtests passed** (2.10s), matching all six CI failures.
- After repair, same four files plus added boundary/security regressions:
  **40 passed, 21 subtests passed, 6 warnings** (1.34s).
- Full `python -m pytest -q --junitxml=...`: **405 passed, 136 subtests passed,
  10 warnings** (31.64s), exit 0.
- Initial full-suite attempt had four missing-file errors because the sparse
  checkout omitted tracked Dasong/inventory/SGIS assets. Materialized those exact
  baseline files, changed no assertions, then ran the successful full suite above.
- `python tools/check_structure.py`: PASS.
- `python -m compileall -q src tests tools`: PASS.
- `python tools/build_dashboard_assets.py`: PASS; generated JS unchanged.
- `node --check dashboard/assets/dashboard.js`: PASS.
- `git diff --check`: PASS.

The ten warnings are existing pandas concat/select_dtypes deprecations, not model
accuracy failures. No skips were used to obtain the passing result.

## Remaining evidence and rollback

Post-push exact SHA, hosted CI run/status, bounded observed CPU pilot and fresh
retraining results are recorded in the existing PR conversation, avoiding a
self-referential commit SHA in this file. Local full-population training, GPU
training, model performance improvement, original-checkpoint parity, browser QA,
main merge and deployment are **NOT_RUN**. Synthetic contract tests are not model
accuracy evidence. The existing CI pilot is development evidence, not new holdout
evidence; its Test must not be used for further model tuning.

Rollback is a normal revert of this repair commit on the same branch. No migration
or data deletion is involved. Reverting restores the known CI failures; it does
not require changing dataset/model artifacts.
