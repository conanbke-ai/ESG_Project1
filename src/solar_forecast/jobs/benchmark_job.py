"""모델별 검증 탐색·독립 하이브리드 선택·최종 Test 산출물 연결."""
from __future__ import annotations

from datetime import datetime, timezone
from copy import deepcopy
import hashlib
import importlib.metadata
import json
from pathlib import Path
import subprocess

import numpy as np
import pandas as pd

from solar_forecast.config_loader import PROJECT_ROOT
from solar_forecast.evaluation.experiment_config import build_candidate_configs, experiment_plan, load_experiment_config
from solar_forecast.infrastructure.artifact_store import create_run_directory, sha256_file, write_json_atomic
from solar_forecast.jobs.training_job import TrainingService
from solar_forecast.models.shared.checkpoint_store import dataset_signature


BENCHMARK_CONTRACT = "solar-optimized-benchmark.v1"
PREDICTION_KEYS = ["plant_id", "timestamp", "forecast_origin", "horizon_hours"]


def align_prediction_frames(frames: dict[str, pd.DataFrame], *, minimum_coverage: float) -> tuple[dict[str, pd.DataFrame], dict]:
    """Align an explicit cohort, failing on changed truth, IDs or low coverage."""
    normalized = {}
    for name, source in frames.items():
        frame = source.copy()
        required = [*PREDICTION_KEYS, "plant", "region", "y_true", "y_pred", "persistence_pred"]
        if missing := set(required) - set(frame):
            raise ValueError(f"{name}: prediction columns missing: {sorted(missing)}")
        frame["timestamp"] = pd.to_datetime(frame["timestamp"], errors="raise")
        frame["forecast_origin"] = pd.to_datetime(frame["forecast_origin"], errors="raise")
        if frame.empty or frame[required].isna().any().any() or frame.duplicated(PREDICTION_KEYS).any():
            raise ValueError(f"{name}: empty, missing or duplicate prediction rows")
        for column in ("y_true", "y_pred", "persistence_pred", "horizon_hours"):
            frame[column] = pd.to_numeric(frame[column], errors="raise")
            if not np.isfinite(frame[column]).all():
                raise ValueError(f"{name}: non-finite {column}")
        expected = frame["forecast_origin"] + pd.to_timedelta(frame["horizon_hours"], unit="h")
        if not expected.eq(frame["timestamp"]).all():
            raise ValueError(f"{name}: target timestamp differs from forecast origin + horizon")
        normalized[name] = frame.set_index(PREDICTION_KEYS).sort_index()
    indices = [frame.index for frame in normalized.values()]
    if not indices:
        raise ValueError("At least one candidate prediction frame is required")
    common = indices[0]
    union = indices[0]
    for index in indices[1:]:
        common = common.intersection(index)
        union = union.union(index)
    coverage = len(common) / len(union) if len(union) else 0
    if not len(common) or coverage < minimum_coverage:
        raise ValueError(f"Common prediction coverage {coverage:.2%} is below {minimum_coverage:.2%}")
    aligned = {name: frame.loc[common].reset_index() for name, frame in normalized.items()}
    reference = next(iter(aligned.values()))
    for name, frame in aligned.items():
        for column in ("plant", "region"):
            if not reference[column].equals(frame[column]):
                raise ValueError(f"{name}: inconsistent {column} on common prediction keys")
        for column in ("y_true", "persistence_pred"):
            if not np.allclose(reference[column], frame[column], rtol=0, atol=1e-7):
                raise ValueError(f"{name}: inconsistent {column} on common prediction keys")
    return aligned, {"common_rows": len(common), "union_rows": len(union), "common_fraction": coverage,
                     "candidate_rows": {name: len(frame) for name, frame in normalized.items()},
                     "dropped_rows": {name: len(frame) - len(common) for name, frame in normalized.items()}}


class BenchmarkService:
    """Execute independently tuned candidates and freeze selection before Test."""

    def __init__(self, *, project_root: Path = PROJECT_ROOT, training_service=None):
        self.project_root = project_root
        self.training_service = training_service or TrainingService()

    def run(self, config_path: Path, *, smoke: bool = False) -> Path:
        from solar_forecast.evaluation.model_selection import BenchmarkModelSelector

        values = load_experiment_config(config_path, project_root=self.project_root)
        if smoke:
            values = deepcopy(values)
            values["horizons_hours"] = [1]
            values["selection_gap_hours"] = 0
            values.setdefault("split", {})["purge_gap_hours"] = 0
            values["optimization_scope"] = "smoke_wiring_only"
            for model, settings in values["models"].items():
                settings["feature_sets"] = settings["feature_sets"][:1]
                if model == "cnn_bilstm":
                    settings["sequence_lengths"] = [24]
        plan = experiment_plan(values, project_root=self.project_root)
        source = Path(plan["input_dataset"])
        if not source.exists() or (source.is_dir() and not any(source.rglob("*.csv")) and not any(source.rglob("*.parquet"))):
            raise FileNotFoundError(f"Observed Gold dataset is unavailable: {source}. Run prepare-data first.")
        output_root = Path(values.get("output_root", "artifacts/benchmarks"))
        output_root = output_root if output_root.is_absolute() else self.project_root / output_root
        run_dir = create_run_directory(output_root)
        provenance = self._provenance(source, values)
        manifest = {"contract": BENCHMARK_CONTRACT, "schema_version": 1,
                    "status": "running", "execution_mode": "smoke" if smoke else "full",
                    "created_at_utc": datetime.now(timezone.utc).isoformat(),
                    "provenance": provenance, "tasks": []}
        write_json_atomic(run_dir / "experiment.json", values)
        write_json_atomic(run_dir / "plan.json", plan)
        write_json_atomic(run_dir / "manifest.json", manifest)
        try:
            for horizon in values["horizons_hours"]:
                candidates = []
                validation_frames = {}
                for candidate_id, config in build_candidate_configs(values, horizon, run_dir, project_root=self.project_root):
                    label = f"{config.model}:{candidate_id}"
                    print(f"Benchmark {horizon}h: {label}", flush=True)
                    candidate_run = self.training_service.run(config, smoke=smoke)
                    result = json.loads((candidate_run / "manifest.json").read_text(encoding="utf-8"))
                    if result.get("status") != "completed":
                        raise ValueError(f"Candidate is not completed: {label}")
                    details = result["details"]
                    if details["evaluation_contract"]["dataset_fingerprint"] != provenance["dataset_fingerprint"]:
                        raise ValueError("Dataset changed during benchmark")
                    write_json_atomic(candidate_run / "resolved_config.json", config.values)
                    candidates.append((label, config, candidate_run, details))
                    # Candidate selection reads Validation alone, never Test metrics.
                    validation_frames[label] = pd.read_csv(details["validation_predictions"], dtype={"plant_id": str})
                aligned, validation_coverage = align_prediction_frames(
                    validation_frames, minimum_coverage=values.get("minimum_common_coverage", 0.95))
                scores = {name: float((frame.y_true - frame.y_pred).abs().mean()) for name, frame in aligned.items()}
                selected = {}
                for model in ("xgboost", "cnn_bilstm"):
                    choices = [entry for entry in candidates if entry[1].model == model]
                    selected[model] = min(choices, key=lambda entry: (scores[entry[0]], entry[0]))
                task_dir = run_dir / f"horizon_{horizon}h"
                optimization = {model: self._candidate_summary(entry, scores[entry[0]]) for model, entry in selected.items()}
                base_decision = {"selection_data": "validation_only", "test_used_for_selection": False,
                                 "candidate_scores": scores, "coverage": validation_coverage,
                                 "selected": optimization}
                # Persist the frozen base choices before accessing Test predictions.
                write_json_atomic(task_dir / "base_selection.json", base_decision)
                contract = self._matching_contract(selected)
                calibration, calibration_coverage = self._merge_selected(selected, "calibration", values)
                test, test_coverage = self._merge_selected(selected, "test", values)
                selector = BenchmarkModelSelector(task_dir,
                    minimum_relative_improvement=float(values.get("minimum_relative_improvement", 0)),
                    selection_gap_hours=max(int(values.get("selection_gap_hours", 0)), horizon))
                result = selector.run(calibration, test, evaluation_contract=contract,
                    provenance={**provenance, "base_selection": str(task_dir / "base_selection.json"),
                                "base_runs": {model: str(entry[2]) for model, entry in selected.items()},
                                "coverage": {"validation": validation_coverage, "calibration": calibration_coverage, "test": test_coverage}})
                manifest["tasks"].append({"horizon_hours": horizon,
                    "selection_path": f"horizon_{horizon}h/selection.json",
                    "selected_model": result["selected_model"], "optimization": optimization,
                    "coverage": {"validation": validation_coverage, "calibration": calibration_coverage, "test": test_coverage}})
                write_json_atomic(run_dir / "manifest.json", manifest)
            if dataset_signature(source) != provenance["dataset_fingerprint"]:
                raise ValueError("Dataset changed during benchmark")
            manifest["status"] = "completed"
            write_json_atomic(run_dir / "manifest.json", manifest)
            return run_dir
        except Exception as exc:
            manifest.update(status="failed", error=f"{type(exc).__name__}: {exc}")
            write_json_atomic(run_dir / "manifest.json", manifest)
            raise

    @staticmethod
    def _candidate_summary(entry: tuple, score: float) -> dict:
        label, config, run_dir, details = entry
        optimizer = details.get("optimizer", {})
        return {"candidate_id": label.split(":", 1)[1], "run_dir": str(run_dir),
                "sequence_length": config.values.get("sequence_length"),
                "feature_columns": config.values["feature_columns"], "validation_mae": score,
                "parameters": optimizer.get("best_params", {}),
                "optimizer": optimizer, "resolved_config": str(run_dir / "resolved_config.json")}

    @staticmethod
    def _matching_contract(selected: dict) -> dict:
        contracts = [entry[3]["evaluation_contract"] for entry in selected.values()]
        for key in ("dataset_fingerprint", "target", "target_unit", "horizon_hours", "prediction_key", "task", "information_set"):
            if key not in contracts[0] or contracts[0].get(key) != contracts[1].get(key):
                raise ValueError(f"Base model evaluation contracts differ or omit {key}")
        return contracts[0]

    @staticmethod
    def _merge_selected(selected: dict, split: str, values: dict) -> tuple[pd.DataFrame, dict]:
        frames = {model: pd.read_csv(entry[3][f"{split}_predictions"], dtype={"plant_id": str}) for model, entry in selected.items()}
        aligned, coverage = align_prediction_frames(frames, minimum_coverage=values.get("minimum_common_coverage", 0.95))
        combined = aligned["xgboost"].drop(columns=["y_pred"]).copy()
        combined["xgb_pred"] = aligned["xgboost"]["y_pred"].to_numpy()
        combined["cnn_pred"] = aligned["cnn_bilstm"]["y_pred"].to_numpy()
        return combined, coverage

    def _provenance(self, source: Path, values: dict) -> dict:
        digest = hashlib.sha256()
        for path in sorted((self.project_root / "src").rglob("*.py")):
            digest.update(path.relative_to(self.project_root).as_posix().encode())
            digest.update(path.read_bytes())
        try:
            revision = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=self.project_root, text=True, stderr=subprocess.DEVNULL).strip()
        except (OSError, subprocess.CalledProcessError):
            revision = None
        runtime = {}
        for package in ("numpy", "pandas", "scikit-learn", "torch", "xgboost", "optuna"):
            try:
                runtime[package] = importlib.metadata.version(package)
            except importlib.metadata.PackageNotFoundError:
                runtime[package] = None
        return {"dataset_fingerprint": dataset_signature(source), "input_dataset": str(source),
                "code_revision": revision, "source_sha256": digest.hexdigest(), "runtime": runtime,
                "experiment_sha256": hashlib.sha256(json.dumps(values, sort_keys=True).encode()).hexdigest(),
                "optimization_scope": values.get("optimization_scope", "configured_search_budget"),
                "data_provenance": values.get("provenance", {}),
                "historical_checkpoint_performance_verified": False}
