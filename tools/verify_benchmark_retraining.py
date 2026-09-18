"""Repeat the observed pilot from fresh optimizer/checkpoint state in a new process.

This checks same-environment training reproducibility. Reusing the existing Test
rows does not create new holdout evidence or reproduce historical checkpoints.
The comparison helpers require only the Python standard library.
"""
from __future__ import annotations

import argparse
from copy import deepcopy
import csv
from datetime import datetime, timedelta, timezone
import gzip
import hashlib
import json
import math
from pathlib import Path
import sys


KEYS = ("plant_id", "timestamp", "forecast_origin", "horizon_hours")
SPLITS = ("validation", "calibration", "test")


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_report(path: Path, report: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


def artifact_path(value: str, run_dir: Path) -> Path:
    """Resolve within this run, including a complete artifact folder relocation."""
    path = Path(value)
    if path.is_absolute():
        if run_dir.name not in path.parts:
            raise ValueError(f"Artifact does not belong to this benchmark: {value}")
        path = Path(*path.parts[path.parts.index(run_dir.name) + 1:])
    resolved = (run_dir / path).resolve()
    if not resolved.is_relative_to(run_dir.resolve()) or not resolved.is_file():
        raise ValueError(f"Missing or external benchmark artifact: {value}")
    return resolved


def _prediction_rows(path: Path, predictions: tuple[str, ...]) -> dict:
    opener = gzip.open if path.suffix == ".gz" else open
    rows = {}
    with opener(path, "rt", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        required = {*KEYS, "plant", "region", "y_true", "persistence_pred", *predictions}
        if not required.issubset(reader.fieldnames or []):
            raise ValueError(f"Prediction columns missing: {path}")
        for row in reader:
            if any(row.get(name) in (None, "") for name in required):
                raise ValueError(f"Missing prediction value: {path}")
            horizon = int(row["horizon_hours"])
            target = datetime.fromisoformat(row["timestamp"])
            origin = datetime.fromisoformat(row["forecast_origin"])
            key = (row["plant_id"], target, origin, horizon)
            if horizon <= 0 or target != origin + timedelta(hours=horizon) or key in rows:
                raise ValueError(f"Duplicate or invalid forecast key: {path}")
            numeric = {name: float(row[name]) for name in ("y_true", "persistence_pred", *predictions)}
            if not all(math.isfinite(value) for value in numeric.values()):
                raise ValueError(f"Non-finite prediction value: {path}")
            rows[key] = {"plant": row["plant"], "region": row["region"], **numeric}
    if not rows:
        raise ValueError(f"Empty prediction artifact: {path}")
    return rows


def compare_predictions(reference: Path, repeated: Path, *, prediction_atol: float = 1e-6,
                        predictions: tuple[str, ...] = ("y_pred",)) -> dict:
    """Require identical forecast keys/truth; tolerate only tiny prediction drift."""
    if not math.isfinite(prediction_atol) or prediction_atol < 0:
        raise ValueError("Prediction tolerance must be finite and nonnegative")
    expected, actual = (_prediction_rows(path, predictions) for path in (reference, repeated))
    if expected.keys() != actual.keys():
        raise ValueError("Forecast keys differ between training runs")
    maxima = dict.fromkeys(predictions, 0.0)
    for key, row in expected.items():
        for name in ("plant", "region", "y_true", "persistence_pred"):
            if row[name] != actual[key][name]:
                raise ValueError(f"{name} differs between training runs at {key}")
        for name in predictions:
            maxima[name] = max(maxima[name], abs(row[name] - actual[key][name]))
    if any(value > prediction_atol for value in maxima.values()):
        raise ValueError(f"Prediction drift exceeds {prediction_atol}: {maxima}")
    return {"rows": len(expected), "prediction_max_abs_delta": maxima,
            "reference_sha256": sha256_file(reference), "repeated_sha256": sha256_file(repeated)}


def fresh_experiment(values: dict, dataset: Path, output: Path) -> dict:
    """Change artifact locations only; keep every model/search/split setting."""
    result = deepcopy(values)
    result["input_dataset"] = str(dataset)
    result["output_root"] = str(output / "runs")
    for settings in result["models"].values():
        overrides = settings.setdefault("training_overrides", {})
        checkpoint = overrides.setdefault("checkpoint", {})
        checkpoint.update(root=str(output / "checkpoints"), resume=False)
        settings.setdefault("optimizer_overrides", {})["storage_path"] = str(output / "optimization/retraining.db")
    return result


def semantic_config(values: dict, *, experiment: bool = False) -> dict:
    """Normalize only the explicit input/output relocations, never model settings."""
    result = deepcopy(values)
    result["input_dataset"] = "<same_observed_dataset_sha256>"
    result["output_root"] = "<isolated_output>"
    containers = list(result["models"].values()) if experiment else [result]
    for container in containers:
        training = container.get("training_overrides", {}) if experiment else container
        if "checkpoint" in training:
            training["checkpoint"]["root"] = "<isolated_checkpoints>"
        optimizer = container.get("optimizer_overrides" if experiment else "optimizer", {})
        if "storage_path" in optimizer:
            optimizer["storage_path"] = "<isolated_optimizer_database>"
    return result


def _candidates(run_dir: Path) -> dict:
    plan = read_json(run_dir / "plan.json")
    expected = {(task["horizon_hours"], item["model"], item["candidate_id"])
                for task in plan["tasks"] for item in task["candidates"]}
    candidates = {}
    for path in (run_dir / "candidates").glob("*/*/*/*/resolved_config.json"):
        config = read_json(path)
        manifest = read_json(path.parent / "manifest.json")
        key = (config["forecast_horizon_hours"], manifest["model"], config["benchmark_candidate_id"])
        if key in candidates or manifest.get("status") != "completed":
            raise ValueError(f"Duplicate or incomplete candidate: {key}")
        candidates[key] = (config, manifest["details"])
    if not expected or candidates.keys() != expected:
        raise ValueError("Candidate artifacts do not match the complete experiment plan")
    return candidates


def _chosen_identity(chosen: dict) -> dict:
    return {key: chosen.get(key) for key in ("candidate_id", "sequence_length", "feature_columns", "parameters")}


def compare_benchmarks(reference: Path, repeated: Path, *, prediction_atol: float = 1e-6) -> dict:
    """Compare every candidate and frozen choice, independent of timestamp folders."""
    manifests = [read_json(path / "manifest.json") for path in (reference, repeated)]
    for manifest in manifests:
        if manifest.get("status") != "completed" or manifest.get("execution_mode") != "full":
            raise ValueError("Retraining comparison requires two completed full benchmarks")
    for field in ("dataset_fingerprint", "source_sha256", "runtime", "optimization_scope"):
        if not manifests[0]["provenance"].get(field) or manifests[0]["provenance"][field] != manifests[1]["provenance"].get(field):
            raise ValueError(f"Benchmark provenance differs or omits {field}")
    configs = [semantic_config(read_json(path / "experiment.json"), experiment=True) for path in (reference, repeated)]
    if configs[0] != configs[1]:
        raise ValueError("Experiment settings changed beyond isolated artifact paths")
    candidate_sets = [_candidates(path) for path in (reference, repeated)]
    if candidate_sets[0].keys() != candidate_sets[1].keys():
        raise ValueError("Candidate sets differ between training runs")
    results = []
    for key, (config, details) in sorted(candidate_sets[0].items()):
        new_config, new_details = candidate_sets[1][key]
        if semantic_config(config) != semantic_config(new_config):
            raise ValueError(f"Resolved model configuration changed: {key}")
        for field in ("evaluation_contract", "temporal_split"):
            if details.get(field) != new_details.get(field):
                raise ValueError(f"Candidate {field} changed: {key}")
        for item, run in ((details, reference), (new_details, repeated)):
            if item.get("checkpoint", {}).get("resumed") or item.get("checkpoint", {}).get("upstream_resumed"):
                raise ValueError(f"Candidate reused a checkpoint: {key}")
            optimizer = item.get("optimizer", {})
            summary = read_json(artifact_path(optimizer["summary_path"], run))
            if summary.get("existing_finished_trials") != 0 or summary.get("executed_trials", 0) < 1:
                raise ValueError(f"Candidate did not execute fresh optimizer trials: {key}")
        if details["optimizer"].get("best_params") != new_details["optimizer"].get("best_params"):
            raise ValueError(f"Selected hyperparameters changed: {key}")
        comparisons = {split: compare_predictions(
            artifact_path(details[f"{split}_predictions"], reference),
            artifact_path(new_details[f"{split}_predictions"], repeated), prediction_atol=prediction_atol,
        ) for split in SPLITS}
        results.append({"horizon_hours": key[0], "model": key[1], "candidate_id": key[2], "splits": comparisons})
    tasks = [{task["horizon_hours"]: task for task in manifest["tasks"]} for manifest in manifests]
    if tasks[0].keys() != tasks[1].keys():
        raise ValueError("Selected horizon sets differ")
    decisions = []
    for horizon, task in sorted(tasks[0].items()):
        other = tasks[1][horizon]
        selection_files = [artifact_path(item["selection_path"], run) for item, run in ((task, reference), (other, repeated))]
        selections = [read_json(path) for path in selection_files]
        identities = []
        for current, selection, selection_file in zip((task, other), selections, selection_files):
            base = read_json(selection_file.parent / "base_selection.json")
            if base.get("test_used_for_selection") is not False:
                raise ValueError("Frozen base selection used Test or omitted its selection contract")
            chosen = {name: _chosen_identity(value) for name, value in current["optimization"].items()}
            if chosen != {name: _chosen_identity(value) for name, value in base["selected"].items()}:
                raise ValueError("Manifest and frozen base selection disagree")
            if selection["selected_model"] != current["selected_model"] or selection["selection_rule"].get("test_used_for_selection") is not False:
                raise ValueError("Frozen selection is inconsistent or used Test")
            rule = {key: value for key, value in selection["selection_rule"].items()
                    if key != "hybrid_relative_improvement"}
            identities.append({"selected_model": current["selected_model"], "base_models": chosen,
                               "selection_rule": rule, "periods": selection.get("periods"),
                               "purge": selection.get("purge")})
        if identities[0] != identities[1]:
            raise ValueError(f"Frozen selected model/candidate/lookback changed at {horizon}h")
        predictions = compare_predictions(
            selection_files[0].parent / selections[0]["files"]["test_predictions"],
            selection_files[1].parent / selections[1]["files"]["test_predictions"],
            prediction_atol=prediction_atol, predictions=("xgb_pred", "cnn_pred", "hybrid_pred", "selected_pred"),
        )
        decisions.append({"horizon_hours": horizon, **identities[0], "test_predictions": predictions})
    return {"status": "passed", "dataset_sha256": manifests[0]["provenance"]["dataset_fingerprint"],
            "source_sha256": manifests[0]["provenance"]["source_sha256"],
            "prediction_atol": prediction_atol, "candidates": results, "selections": decisions}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--pilot-dir", type=Path, default=Path("artifacts/verification/observed_benchmark"))
    parser.add_argument("--output-dir", type=Path, default=None)
    args = parser.parse_args()
    root = args.project_root.resolve()
    pilot = (args.pilot_dir if args.pilot_dir.is_absolute() else root / args.pilot_dir).resolve()
    output = args.output_dir or pilot / "retraining"
    output = (output if output.is_absolute() else root / output).resolve()
    if output.exists() and any(output.iterdir()):
        raise ValueError("Retraining output must be absent or empty; existing evidence is never overwritten")
    output.mkdir(parents=True, exist_ok=True)
    report_path = output / "retraining_report.json"
    report = {"contract": "solar-benchmark-retraining-reproducibility.v1", "status": "started",
              "created_at_utc": datetime.now(timezone.utc).isoformat(), "training_performed": False,
              "new_holdout_evidence": False, "historical_checkpoint_performance_verified": False,
              "scope": "same_environment_fresh_training_on_the_existing_observed_pilot_and_Test",
              "prediction_atol": 1e-6}
    write_report(report_path, report)
    try:
        from run_observed_benchmark_pilot import configure_cpu_runtime, runtime_versions
        configure_cpu_runtime()
        initial = read_json(pilot / "observed_pilot_report.json")
        if initial.get("status") != "completed":
            raise ValueError("A successfully completed observed pilot is required")
        dataset = pilot / "observed_plant_year.csv"
        digest = sha256_file(dataset)
        if digest != initial["provenance"]["input_dataset_sha256"]:
            raise ValueError("Observed pilot input SHA256 changed")
        reference = Path(initial["benchmark_run_dir"])
        if not reference.is_absolute():
            reference = root / reference
        if not reference.is_dir():
            reference = root / "artifacts/benchmarks" / reference.name
        manifest = read_json(reference / "manifest.json")
        values = read_json(pilot / "pilot_experiment.json")
        if values != read_json(reference / "experiment.json"):
            raise ValueError("Pilot configuration differs from the successful first run")
        sys.path.insert(0, str(root / "src"))
        import torch
        torch.set_num_threads(1)
        torch.set_num_interop_threads(1)
        runtime = runtime_versions()
        report["runtime"] = runtime
        if runtime != initial.get("runtime"):
            raise ValueError("First pilot and retraining runtime evidence differ, including CPU/backend settings")
        from solar_forecast.jobs.benchmark_job import BenchmarkService
        config = fresh_experiment(values, dataset, output)
        current = BenchmarkService(project_root=root)._provenance(dataset, config)
        for field in ("dataset_fingerprint", "source_sha256", "runtime"):
            if current[field] != manifest["provenance"][field]:
                raise ValueError(f"First pilot and current environment differ: {field}")
        config_path = output / "retraining_experiment.json"
        write_report(config_path, config)
        report.update(reference_run=str(reference), dataset_sha256=digest,
                      fresh_checkpoint_root=str(output / "checkpoints"),
                      fresh_optimizer_database=str(output / "optimization/retraining.db"), training_performed=True)
        write_report(report_path, report)
        repeated = BenchmarkService(project_root=root).run(config_path, smoke=False)
        report["repeated_run"] = str(repeated)
        report.update(compare_benchmarks(reference, repeated))
        if sha256_file(dataset) != digest:
            raise ValueError("Observed input changed during retraining")
    except Exception as exc:
        report.update(status="failed", error=f"{type(exc).__name__}: {exc}")
    write_report(report_path, report)
    print(json.dumps({"status": report["status"], "report": str(report_path),
                      "error": report.get("error")}, ensure_ascii=False))
    raise SystemExit(0 if report["status"] == "passed" else 1)


if __name__ == "__main__":
    main()
