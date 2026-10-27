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

## Follow-up: previously skipped integration gates

The first repair was published as `4c1d5bc7e882619c27471e05d1e0c6d3235d0cde`.
[CI #40438890866](https://github.com/conanbke-ai/ESG_Project1/actions/runs/40438890866)
passed structure and all 405 regression tests, then exposed an observed-pilot
failure before training: `Temporal split produced empty partitions ['test']`.

The pilot selected a complete 2024 plant-year using its availability audit's
fraction-derived boundaries, but executed the inherited production calendar whose
Test is in 2025. The preflight and trainer were therefore describing different
cohorts. The pilot now freezes its already-audited boundaries as explicit shared
calendar cutoffs for every candidate. Requested production dates and effective
pilot dates are both recorded in `pilot_overrides`. `optimized.json`, production
calendar enforcement and data eligibility remain unchanged. No model score is
used to select these dates.

The observed-only pilot also explicitly disables the inherited archived-weather
candidates, which require separate local archives not included in this pilot's
retained inputs. Requested archive configuration remains in its evidence. This is
not a fallback in production forecasting, a weather-availability waiver, or
validation of the full production experiment. Its budget remains two XGBoost
trials / one CNN trial, three CNN epochs, 1h/24h horizons and 24h/168h lookbacks.

Running that pilot locally then exposed a stored-model verifier error: CNN
capacity-factor output was compared directly with exported MWh predictions.
The verifier now restores the checkpoint's target-transform/capacity-column
contract and each target row's float32 capacity scale. Legacy checkpoints without
the field retain identity/MWh semantics. Contract mismatches and missing capacity
fail closed; invalid target capacity is excluded consistently with training.
Saved model weights, exported predictions, training code and `1e-6` tolerance
remain unchanged. Tests include actual checkpoint serialization/export, reordered
prediction keys, varying capacity, non-feature capacity passthrough and no refit.

### Executed follow-up evidence

- Pilot/calendar and retraining unit tests: **19 passed, 13 subtests passed**.
- CNN replay/partition tests: **9 passed, 4 warnings**.
- Full regression: **410 passed, 136 subtests passed, 14 warnings**, 33.44s, exit 0.
- Generator, structure, compileall and diff whitespace: PASS.
- Local retained-cohort CPU integration: all six candidate searches and four
  selected base fits completed. Input: `ewp:광양항광양냉장태양광`, 2024, 8,952 rows
  including 168 observed context hours; dataset SHA-256
  `866f22e18656a4a18fa30c9433710318792c3503160eb77e07e9c74ccaf68fee`.
  This exact CSV came from CI #40438890866 artifact 12705053243; local Gold
  preparation was not rerun or substituted with synthetic data.
- Stored-model replay: **4/4 passed**, 1,175 Test rows per selected model/horizon.
  Maximum XGBoost delta: `1.4871597286969518e-08` MWh; maximum CNN delta:
  `9.974659986866641e-17` MWh, within unchanged tolerance `1e-6`.
- `python tools/verify_benchmark_retraining.py --pilot-dir <retained-cohort-run>`:
  fresh optimizer/checkpoint roots in a separate process, **passed**, all six
  search candidates and both final selections compared. Final XGBoost/CNN/Hybrid/
  selected prediction deltas were **0.0** for both horizons in this environment.

This confirms verifier/runner consistency for the bounded CPU experiment. It does
not establish original-model equivalence, model improvement, unseen holdout
generalization, production readiness or identical results on a different CPU/GPU.
The repeated Test is existing development evidence, not a new test population.
Final follow-up commit and hosted CI outcome are in the PR conversation.
