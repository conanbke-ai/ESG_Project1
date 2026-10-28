"""Compare two checkouts using frozen data and isolated, small CPU training runs.

This checks refactor parity, not historical production forecasting accuracy.
No API calls, hyperparameter search, production checkpoints, or installs occur.
"""
from __future__ import annotations

import argparse
import csv
from datetime import datetime, timedelta, timezone
import hashlib
import importlib
import importlib.metadata
import json
import math
import os
from pathlib import Path
import platform
import random
import subprocess
import sys


MODELS = ("xgboost", "cnn_bilstm")
SPLITS = ("validation", "calibration", "test")
SEED = 42


def write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2, default=str) + "\n", encoding="utf-8")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def create_output_directory(path: Path) -> None:
    if path.exists() and (not path.is_dir() or any(path.iterdir())):
        raise ValueError("Output directory must be absent or empty; existing evidence is never overwritten")
    path.mkdir(parents=True, exist_ok=True)


def create_synthetic_fixture(path: Path, feature_columns: list[str]) -> None:
    """Write two fictional plants with hourly continuity and explicit missingness."""
    records = []
    start = datetime(2024, 6, 1)
    for step in range(480):
        stamp = start + timedelta(hours=step)
        daylight = max(0.0, math.sin(math.pi * (stamp.hour - 6) / 12))
        for entity in range(2):
            capacity = 1.0 + entity * 0.5
            target = capacity * daylight * (0.8 + 0.1 * math.sin(step / 53))
            values = {
                "temperature_c": 20 + 8 * daylight if step % 37 else "",
                "precipitation_mm": 0.2 if step % 97 == 0 else 0.0,
                "sunshine_hours": daylight, "solar_irradiance_mj_m2": daylight * 2,
                "wind_speed_mps": 2 + math.sin(step / 19), "humidity_pct": 65 - 20 * daylight,
                "total_cloud_cover_tenths": 3.0, "low_mid_cloud_cover_tenths": 2.0,
                "hour": stamp.hour, "dayofweek": stamp.weekday(), "month": stamp.month,
                "hour_sin": math.sin(2 * math.pi * stamp.hour / 24),
                "hour_cos": math.cos(2 * math.pi * stamp.hour / 24),
                "dayofyear_sin": math.sin(2 * math.pi * stamp.timetuple().tm_yday / 366),
                "dayofyear_cos": math.cos(2 * math.pi * stamp.timetuple().tm_yday / 366),
                "generation_lag_24h_mwh": capacity * daylight * (0.8 + 0.1 * math.sin((step - 24) / 53)) if step >= 24 else "",
                "generation_lag_168h_mwh": capacity * daylight * (0.8 + 0.1 * math.sin((step - 168) / 53)) if step >= 168 else "",
                "generation_rolling_7d_mean_mwh": capacity * 0.25 if step >= 168 else "",
                "capacity_mw": capacity, "tilt_deg": "", "station_latitude": 35 + entity,
                "station_longitude": 127 + entity, "station_elevation_m": 40 + entity * 20,
                "solar_elevation_sin": daylight, "clear_sky_irradiance_proxy": daylight,
                "is_daylight": int(daylight > 1e-10),
            }
            missing = set(feature_columns) - values.keys()
            if missing:
                raise ValueError(f"Synthetic fixture does not define configured features: {sorted(missing)}")
            records.append({
                "timestamp": stamp.isoformat(sep=" "), "plant_id": f"synthetic-{entity}",
                "region": f"synthetic-region-{entity}", "plant": f"synthetic-plant-{entity}",
                "energy_source": "solar", "quality_train_eligible": True,
                "generation_mwh": target, **{name: values[name] for name in feature_columns},
            })
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(records[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(records)


def experiment_config(original: dict, dataset: Path, output: Path) -> dict:
    values = json.loads(json.dumps(original))
    values.update({
        "input_dataset": str(dataset), "seed": SEED, "optimizer": {"enabled": False},
        "use_optuna": False, "epochs": 2, "sequence_length": 24, "batch_size": 64,
        "shuffle_training_batches": True, "n_estimators": 20, "n_jobs": 1,
        "purge_gap_hours": 24,
        "checkpoint": {"enabled": False, "resume": False, "root": str(output / "disabled_checkpoints")},
    })
    return values


def revision(root: Path) -> str:
    result = subprocess.run(["git", "-C", str(root), "rev-parse", "HEAD"], capture_output=True, text=True)
    return result.stdout.strip() if result.returncode == 0 else "unavailable"


def source_identity(root: Path) -> dict:
    files = sorted([*(root / "src").rglob("*.py"), *(root / "config" / "models").glob("*.json")])
    digest = hashlib.sha256()
    for path in files:
        name = path.relative_to(root).as_posix().encode("utf-8")
        content = path.read_bytes()
        digest.update(len(name).to_bytes(8, "big") + name)
        digest.update(len(content).to_bytes(8, "big") + content)
    status = subprocess.run(["git", "-C", str(root), "status", "--porcelain", "--untracked-files=normal"], capture_output=True, text=True)
    return {"root": str(root), "revision": revision(root), "git_dirty": bool(status.stdout.strip()) if status.returncode == 0 else "unavailable", "source_sha256": digest.hexdigest(), "source_files": len(files)}


def runtime_versions() -> dict:
    packages = {}
    for name in ("numpy", "pandas", "scikit-learn", "torch", "xgboost", "optuna"):
        try:
            packages[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            packages[name] = "unavailable"
    return {"python": sys.version, "platform": platform.platform(), "packages": packages}


def check_artifact_compatibility(model: str, baseline: dict, candidate: dict) -> dict:
    """Candidate runtime loads baseline artifacts; inference is checked separately."""
    import torch

    checks = {"baseline_loaded_with_candidate": True, "features_equal": baseline["features"] == candidate["features"]}
    if model == "cnn_bilstm":
        from solar_forecast.models.cnn_bilstm.evaluation import load_checkpoint
        load_checkpoint(baseline["checkpoint_path"], device=torch.device("cpu"))
        before = torch.load(baseline["checkpoint_path"], map_location="cpu", weights_only=False)
        after = torch.load(candidate["checkpoint_path"], map_location="cpu", weights_only=False)
        checks.update({
            "network_config_equal": before["config"] == after["config"],
            "preprocessing_equal": before["preprocessing"] == after["preprocessing"],
            "feature_order_equal": before["feature_columns"] == after["feature_columns"],
            "state_keys_equal": before["model_state"].keys() == after["model_state"].keys(),
        })
        maximum = 0.0
        for key, tensor in before["model_state"].items():
            other = after["model_state"][key]
            torch.testing.assert_close(tensor, other, atol=1e-6, rtol=1e-6)
            maximum = max(maximum, float((tensor.to(torch.float64) - other.to(torch.float64)).abs().max()))
        checks["maximum_state_difference"] = maximum
        checks["preprocessing_check"] = "saved training-split preprocessing equality; serving transform not independently tested"
    else:
        from xgboost import XGBRegressor
        before, after = XGBRegressor(), XGBRegressor()
        before.load_model(baseline["model_path"])
        after.load_model(candidate["model_path"])
        checks["tree_dump_equal"] = before.get_booster().get_dump(dump_format="json") == after.get_booster().get_dump(dump_format="json")
        checks["feature_names_equal"] = before.get_booster().feature_names == after.get_booster().feature_names
        baseline_preprocessing = json.loads(Path(baseline["preprocessing_path"]).read_text(encoding="utf-8"))
        candidate_preprocessing = json.loads(Path(candidate["preprocessing_path"]).read_text(encoding="utf-8"))
        for key in ("strategy", "missing_value", "feature_missing_fraction", "temporal_split"):
            checks[f"{key}_equal"] = baseline_preprocessing[key] == candidate_preprocessing[key]
    checks["status"] = "passed" if all(value for value in checks.values() if isinstance(value, bool)) else "failed"
    return checks


def run_worker(request_path: Path) -> None:
    request = json.loads(request_path.read_text(encoding="utf-8"))
    root = Path(request["root"])
    sys.path.insert(0, str(root / "src"))
    import numpy as np
    import torch

    random.seed(SEED)
    np.random.seed(SEED)
    torch.manual_seed(SEED)
    torch.set_num_threads(1)
    torch.set_num_interop_threads(1)
    torch.use_deterministic_algorithms(True)
    torch.backends.mkldnn.enabled = False
    if torch.cuda.is_available():
        raise RuntimeError("Parity worker requires CPU-only execution")
    candidate = request["side"] == "candidate"
    settings = importlib.import_module("solar_forecast.config_loader" if candidate else "solar_forecast.settings")
    model = request["model"]
    module_name = f"solar_forecast.models.{model}" + (".trainer" if candidate else "")
    trainer_type = getattr(importlib.import_module(module_name), "XGBoostTrainer" if model == "xgboost" else "CnnBiLstmTrainer")
    config = settings.ModelJobConfig(model, request["values"].get("profile", "optimized"), request["values"], Path(request["config_source"]))
    run_dir = Path(request["run_dir"])
    run_dir.mkdir(parents=True, exist_ok=False)
    # Imports can have independent RNG side effects in differently arranged
    # packages. Reset immediately before the identical training entry point.
    random.seed(SEED)
    np.random.seed(SEED)
    torch.manual_seed(SEED)
    result = trainer_type().train(config, run_dir, smoke=False)
    result["runtime"] = runtime_versions()
    if candidate:
        baseline = json.loads(Path(request["baseline_result"]).read_text(encoding="utf-8"))
        result["artifact_compatibility"] = check_artifact_compatibility(model, baseline, result)
    write_json(Path(request["result_path"]), result)


def execute_worker(request: dict, request_path: Path) -> dict:
    write_json(request_path, request)
    environment = os.environ.copy()
    environment.update({
        "PYTHONPATH": str(Path(request["root"]) / "src"), "PYTHONHASHSEED": str(SEED),
        "PYTHONDONTWRITEBYTECODE": "1", "CUDA_VISIBLE_DEVICES": "",
        "OMP_NUM_THREADS": "1", "MKL_NUM_THREADS": "1", "OPENBLAS_NUM_THREADS": "1",
        "NUMEXPR_NUM_THREADS": "1",
    })
    log_path = request_path.with_suffix(".log")
    with log_path.open("w", encoding="utf-8") as stream:
        result = subprocess.run(
            [sys.executable, str(Path(__file__).resolve()), "--worker", str(request_path)],
            cwd=request["root"], env=environment, stdout=stream, stderr=subprocess.STDOUT,
        )
    if result.returncode:
        raise RuntimeError(f"{request['side']} {request['model']} worker failed; inspect {log_path}")
    return json.loads(Path(request["result_path"]).read_text(encoding="utf-8"))


def run_comparison(args: argparse.Namespace) -> dict:
    baseline, candidate, output = (Path(value).resolve() for value in (args.baseline_root, args.candidate_root, args.output_dir))
    for root in (baseline, candidate):
        if not (root / "src" / "solar_forecast").is_dir():
            raise ValueError(f"Checkout source is missing: {root}")
    create_output_directory(output)
    report = {
        "contract": "solar-model-reorganization-parity.v1", "status": "running",
        "mode": "real_dataset_small_training_parity" if args.dataset else "synthetic_small_training_parity",
        "production_accuracy_verified": False, "historical_performance_verified": False,
        "limitations": [
            "Two CNN epochs and twenty XGBoost trees with optimization disabled check refactoring, not production model quality.",
            "Current CNN emits one value after its observation window; horizon_hours=24 metadata does not verify a simultaneous 24-hour forecast.",
            "Synthetic inputs represent fictional plants and are never admitted into Gold.",
            "Artifact loading and saved preprocessing equality do not independently validate a serving preprocessing implementation.",
        ],
        "started_at_utc": datetime.now(timezone.utc).isoformat(),
        "checkouts": {name: source_identity(root) for name, root in (("baseline", baseline), ("candidate", candidate))},
        "seed": SEED, "cpu_threads": 1, "configurations": {}, "models": {},
    }
    try:
        sources = {side: {model: root / "config" / "models" / f"{model}.json" for model in MODELS} for side, root in (("baseline", baseline), ("candidate", candidate))}
        originals = {side: {model: json.loads(path.read_text(encoding="utf-8")) for model, path in paths.items()} for side, paths in sources.items()}
        dataset = Path(args.dataset).resolve() if args.dataset else output / "synthetic_fixture.csv"
        if args.dataset and not Path(args.dataset).is_absolute():
            raise ValueError("--dataset must be an absolute path to one frozen CSV file")
        if not args.dataset:
            features = list(dict.fromkeys(name for values in originals["candidate"].values() for name in values["feature_columns"]))
            create_synthetic_fixture(dataset, features)
        if not dataset.is_file() or dataset.suffix.lower() != ".csv":
            raise ValueError("Dataset must be an existing frozen CSV file")
        dataset_hash = sha256_file(dataset)
        report["dataset"] = {"path": str(dataset), "sha256": dataset_hash, "bytes": dataset.stat().st_size, "synthetic": not bool(args.dataset)}
        sys.path.insert(0, str(candidate / "src"))
        from solar_forecast.evaluation.model_parity import compare_prediction_files
        for model in MODELS:
            results = {}
            configurations = {side: experiment_config(originals[side][model], dataset, output) for side in sources}
            if configurations["baseline"] != configurations["candidate"]:
                raise ValueError(f"Effective {model} configurations differ; resolve that difference before claiming code-only parity")
            report["configurations"][model] = {}
            for side, root in (("baseline", baseline), ("candidate", candidate)):
                config_path = output / f"{side}_{model}_config.json"
                write_json(config_path, configurations[side])
                report["configurations"][model][side] = {"original_sha256": sha256_file(sources[side][model]), "effective_sha256": sha256_file(config_path), "effective_path": str(config_path)}
                if sha256_file(dataset) != dataset_hash:
                    raise ValueError("Frozen dataset changed during the comparison")
                request = {
                    "side": side, "root": str(root), "model": model, "values": configurations[side],
                    "config_source": str(sources[side][model]), "run_dir": str(output / f"{side}_{model}"),
                    "result_path": str(output / f"{side}_{model}_result.json"),
                    "baseline_result": str(output / f"baseline_{model}_result.json"),
                }
                results[side] = execute_worker(request, output / f"{side}_{model}_request.json")
            comparisons = {split: compare_prediction_files(Path(results["baseline"][f"{split}_predictions"]), Path(results["candidate"][f"{split}_predictions"])) for split in SPLITS}
            compatibility = results["candidate"]["artifact_compatibility"]
            equal_runtime = results["baseline"]["runtime"] == results["candidate"]["runtime"]
            equal_split = results["baseline"]["temporal_split"] == results["candidate"]["temporal_split"]
            report["models"][model] = {"status": "passed" if all(result["status"] == "passed" for result in comparisons.values()) and compatibility["status"] == "passed" and equal_runtime and equal_split else "failed", "predictions": comparisons, "artifact_compatibility": compatibility, "runtime_equal": equal_runtime, "runtime": results["candidate"]["runtime"], "temporal_split_equal": equal_split}
        if sha256_file(dataset) != dataset_hash:
            raise ValueError("Frozen dataset changed during the comparison")
        for side, root in (("baseline", baseline), ("candidate", candidate)):
            if source_identity(root)["source_sha256"] != report["checkouts"][side]["source_sha256"]:
                raise ValueError(f"{side} source or model configuration changed during the comparison")
        report["status"] = "passed" if all(model["status"] == "passed" for model in report["models"].values()) else "failed"
    except Exception as exc:
        report.update(status="failed", error=f"{type(exc).__name__}: {exc}")
    report["report_path"] = str(output / "report.json")
    write_json(output / "report.json", report)
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-root")
    parser.add_argument("--candidate-root", default=str(Path(__file__).resolve().parents[1]))
    parser.add_argument("--output-dir")
    parser.add_argument("--dataset", help="Absolute path to a frozen real CSV; small training budget still applies")
    parser.add_argument("--worker", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.worker:
        run_worker(Path(args.worker))
        return 0
    if not args.baseline_root or not args.output_dir:
        parser.error("--baseline-root and --output-dir are required")
    report = run_comparison(args)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
