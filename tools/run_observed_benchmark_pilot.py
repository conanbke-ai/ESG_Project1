"""Run a bounded benchmark on a measured plant-year from retained public archives.

This is a real-data integration pilot, not full-population optimization or an
original-checkpoint reproduction. No measured targets are interpolated or made up.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import importlib.metadata
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys


OPTIMIZATION_SCOPE = "bounded_observed_pilot"
DATA_ORIGIN = "retained_official_observations"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(
        json.dumps(value, ensure_ascii=False, indent=2, default=str) + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)


def continuous_evaluation_coverage(cohort, split: dict, *, minimum_coverage: float = 0.95) -> dict:
    """Check shared evaluation windows before selecting a data-availability cohort."""
    import numpy as np
    import pandas as pd
    from solar_forecast.evaluation.forecast_samples import forecast_window_positions
    from solar_forecast.evaluation.temporal_split import TemporalSplitConfig, TemporalSplitter

    splitter = TemporalSplitter(TemporalSplitConfig(
        validation_fraction=float(split.get("validation_fraction", 0.15)),
        calibration_fraction=float(split.get("calibration_fraction", 0.10)),
        test_fraction=float(split.get("test_fraction", 0.15)),
        gap_hours=max(int(split.get("purge_gap_hours", 168)), 24),
    ))
    times = cohort["timestamp"].reset_index(drop=True)
    boundaries = splitter.boundaries(times)
    labels = splitter.labels(times, boundaries)
    coverage = {"sequence_length_hours": 168, "boundaries": boundaries.to_dict(), "horizons": {}}
    qualified = True
    for horizon in (1, 24):
        tabular, _ = forecast_window_positions(times, horizon_hours=horizon, sequence_length=1)
        sequence, _ = forecast_window_positions(times, horizon_hours=horizon, sequence_length=168)
        task = {}
        for name in ("validation", "calibration", "test"):
            positions = np.flatnonzero(labels.eq(name).fillna(False).to_numpy(dtype=bool))
            available = np.intersect1d(tabular, positions)
            common = np.intersect1d(sequence, positions)
            fraction = len(common) / len(available) if len(available) else 0.0
            task[name] = {"tabular_rows": len(available), "common_rows": len(common), "common_fraction": fraction}
            qualified &= len(common) >= 32 and fraction >= minimum_coverage
        # The independent gate-fitting and selection blocks also need enough
        # measured timestamps after their internal chronological purge.
        calibration_positions = np.intersect1d(
            sequence, np.flatnonzero(labels.eq("calibration").fillna(False).to_numpy(dtype=bool))
        )
        calibration_times = times.iloc[calibration_positions]
        if len(calibration_times) >= 4:
            midpoint = calibration_times.iloc[len(calibration_times) // 2]
            gate_rows = int(calibration_times.lt(midpoint - pd.Timedelta(hours=max(168, horizon))).sum())
            selection_rows = int(calibration_times.ge(midpoint).sum())
        else:
            gate_rows = selection_rows = 0
        task["calibration"]["gate_fit_rows_after_purge"] = gate_rows
        task["calibration"]["selection_rows"] = selection_rows
        qualified &= gate_rows >= 2 and selection_rows >= 2
        coverage["horizons"][str(horizon)] = task
    coverage["qualified"] = bool(qualified)
    return coverage


def select_observed_plant_year(frame, *, split: dict | None = None, minimum_coverage: float = 0.95):
    """Select solely by calendar coverage, before looking at any model score."""
    import numpy as np
    import pandas as pd

    required = {
        "timestamp", "plant_id", "energy_source", "quality_train_eligible",
        "generation_mwh", "temperature_c", "wind_speed_mps", "humidity_pct",
    }
    missing = required.difference(frame.columns)
    if missing:
        raise ValueError(f"Gold pilot columns are missing: {sorted(missing)}")
    eligible = frame["quality_train_eligible"].astype(str).str.lower().isin(["true", "1"])
    observed = frame.loc[eligible & frame["energy_source"].eq("solar")].copy()
    observed["timestamp"] = pd.to_datetime(observed["timestamp"], errors="raise")
    target = pd.to_numeric(observed["generation_mwh"], errors="coerce")
    if observed.empty or not (np.isfinite(target) & target.ge(0)).all():
        raise ValueError("Eligible solar Gold must contain finite, nonnegative measured targets")
    if observed.duplicated(["plant_id", "timestamp"]).any():
        raise ValueError("Gold contains duplicate plant-hour observations")
    if not observed["timestamp"].eq(observed["timestamp"].dt.floor("h")).all():
        raise ValueError("Gold contains timestamps outside the hourly grid")

    candidates = []
    rejected_continuity = 0
    for (plant, year), group in observed.groupby(
        ["plant_id", observed["timestamp"].dt.year], sort=True
    ):
        start = pd.Timestamp(year=int(year), month=1, day=1)
        end = pd.Timestamp(year=int(year) + 1, month=1, day=1)
        expected_hours = int((end - start) / pd.Timedelta(hours=1))
        coverage = len(group) / expected_hours
        weather_coverage = float(
            group[["temperature_c", "wind_speed_mps", "humidity_pct"]].notna().all(axis=1).mean()
        )
        if (
            coverage >= 0.95
            and group["timestamp"].max() - group["timestamp"].min() >= pd.Timedelta(days=364)
            and weather_coverage >= 0.90
        ):
            candidate = observed.loc[
                observed["plant_id"].eq(plant)
                & observed["timestamp"].ge(start - pd.Timedelta(hours=168))
                & observed["timestamp"].lt(end)
            ].sort_values("timestamp", kind="stable").reset_index(drop=True)
            continuous = continuous_evaluation_coverage(candidate, split or {}, minimum_coverage=minimum_coverage)
            if not continuous["qualified"]:
                rejected_continuity += 1
                continue
            candidates.append({
                "plant_id": str(plant), "calendar_year": int(year),
                "observed_hours": len(group), "expected_hours": expected_hours,
                "hourly_coverage": coverage, "core_weather_coverage": weather_coverage,
                "continuous_window_coverage": continuous,
            })
    if not candidates:
        raise ValueError("No measured solar plant-year meets hourly, weather, and continuous evaluation-window coverage")
    candidates.sort(key=lambda row: (-row["calendar_year"], -row["hourly_coverage"], row["plant_id"]))
    selected = candidates[0]
    start = pd.Timestamp(year=selected["calendar_year"], month=1, day=1)
    end = pd.Timestamp(year=selected["calendar_year"] + 1, month=1, day=1)
    # Preserve existing prior context only; missing timestamps are never created.
    cohort = observed.loc[
        observed["plant_id"].astype(str).eq(selected["plant_id"])
        & observed["timestamp"].ge(start - pd.Timedelta(hours=168))
        & observed["timestamp"].lt(end)
    ].sort_values("timestamp", kind="stable").reset_index(drop=True)
    selected.update({
        "selection_policy": "newest qualifying calendar year; hourly coverage descending; plant_id ascending",
        "model_performance_used_for_selection": False,
        "qualifying_plant_years": len(candidates),
        "plant_years_rejected_for_window_continuity": rejected_continuity,
        "context_rows_before_calendar_year": int(cohort["timestamp"].lt(start).sum()),
        "input_rows": len(cohort),
        "input_start": cohort["timestamp"].min().isoformat(),
        "input_end": cohort["timestamp"].max().isoformat(),
        "target_interpolation": "none",
        "missing_hour_policy": "retain observed rows only; forecast windows enforce hourly continuity",
    })
    return cohort, selected


def prepare_official_gold(project_root: Path, destination: Path) -> Path:
    subprocess.run(
        [
            sys.executable, str(project_root / "app.py"), "prepare-data",
            "--input-root", str(project_root / "file/solar_data_file"),
            "--weather-root", str(project_root / "file/KMA_data_file"),
            "--merged-source", str(project_root / "file/merge_data/val.csv"),
            "--output-dir", str(destination), "--no-collected-downloads",
        ],
        cwd=project_root,
        check=True,
    )
    return destination / "model_ready.csv.gz"


def capture_input_provenance(project_root: Path, prepared: Path, output: Path) -> dict:
    manifest_names = ("model_ready_manifest.json", "generation_manifest.json", "plant_registry.csv")
    saved = []
    for name in manifest_names:
        source = prepared / name
        destination = output / "provenance" / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, destination)
        saved.append({"path": str(source), "saved_copy": str(destination), "sha256": sha256_file(source)})
    # The generation manifest already records every admitted raw-file hash.
    # ASOS source hashes are recorded here because Gold's manifest records years.
    source_paths = [
        *sorted((project_root / "file/KMA_data_file").glob("*.csv")),
        *sorted((project_root / "file/solar_data_file/location").glob("*.csv")),
        project_root / "file/merge_data/val.csv",
        project_root / "config/reviewed_weather_mappings.json",
    ]
    return {
        "upstream_artifacts": saved,
        "weather_and_mapping_sources": [
            {"path": str(path.relative_to(project_root)), "bytes": path.stat().st_size, "sha256": sha256_file(path)}
            for path in source_paths
        ],
    }


def configure_cpu_runtime() -> None:
    # Set these before importing NumPy, Torch, or the benchmark service.
    for name in ("OMP_NUM_THREADS", "MKL_NUM_THREADS", "OPENBLAS_NUM_THREADS", "NUMEXPR_NUM_THREADS"):
        os.environ[name] = "1"
    os.environ["CUDA_VISIBLE_DEVICES"] = ""


def runtime_versions() -> dict:
    packages = {}
    for name in ("numpy", "pandas", "scikit-learn", "torch", "xgboost", "optuna"):
        try:
            packages[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            packages[name] = "unavailable"
    return {"python": sys.version, "packages": packages}


def build_pilot_config(base: dict, dataset: Path, output: Path, project_root: Path, provenance: dict) -> dict:
    """Bound resources explicitly while running ordinary, non-smoke trainers."""
    values = json.loads(json.dumps(base))
    values.update({
        "name": "observed_plant_year_benchmark_pilot",
        "description": "공식 실측 발전소 1개 연도의 제한된 탐색; 전체 모집단 최적화 결과가 아님",
        "input_dataset": str(dataset),
        "output_root": str(project_root / "artifacts/benchmarks"),
        "horizons_hours": [1, 24],
        "seed": 42,
        "optimization_scope": OPTIMIZATION_SCOPE,
        "data_origin": DATA_ORIGIN,
        "provenance": provenance,
    })
    checkpoint = {"enabled": True, "resume": False, "root": str(output / "checkpoints")}
    for name, settings in values["models"].items():
        settings["feature_sets"] = [settings["feature_sets"][0]]
        settings["optimize"] = True
        settings["max_trials_per_candidate"] = 1 if name == "cnn_bilstm" else 2
        settings["timeout_seconds_per_candidate"] = 120
        settings["training_overrides"] = {
            **settings.get("training_overrides", {}),
            "n_jobs": 1, "checkpoint": checkpoint,
        }
        settings["optimizer_overrides"] = {
            **settings.get("optimizer_overrides", {}),
            "storage_path": str(output / "optimization/pilot.db"),
            "startup_trials": 1, "pruner_startup_trials": 1,
        }
    cnn = values["models"]["cnn_bilstm"]
    cnn["sequence_lengths"] = [24, 168]
    cnn["training_overrides"].update({"epochs": 3, "batch_size": 128, "early_stopping_patience": 3})
    cnn["optimizer_overrides"].update({
        "trial_epochs": 3, "early_stopping_patience": 3,
        "tuning_train_max_sequences": 10000, "tuning_validation_max_sequences": 10000,
        "search_space": {
            name: {"type": "categorical", "choices": [value]}
            for name, value in {
                "cnn_channels": 16, "kernel_size": 3, "lstm_hidden": 32,
                "lstm_layers": 1, "dense_units": 32, "dropout": 0.1,
                "lr": 0.001, "weight_decay": 0.0001,
            }.items()
        },
    })
    xgb = values["models"]["xgboost"]
    xgb["training_overrides"].update({"n_estimators": 128, "early_stopping_rounds": 10})
    xgb["optimizer_overrides"].update({
        "trial_max_estimators": 128, "early_stopping_rounds": 10,
        "tuning_train_max_rows": 10000, "tuning_validation_max_rows": 10000,
        "search_space": {
            "max_depth": {"type": "categorical", "choices": [3, 5]},
            "learning_rate": {"type": "categorical", "choices": [0.03, 0.08]},
            **{
                name: {"type": "categorical", "choices": [value]}
                for name, value in {
                    "min_child_weight": 1.0, "subsample": 0.9, "colsample_bytree": 0.9,
                    "reg_alpha": 0.0, "reg_lambda": 1.0, "gamma": 0.0, "max_bin": 128,
                }.items()
            },
        },
    })
    return values


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--config", type=Path, default=Path("config/experiments/optimized.json"))
    parser.add_argument("--output-dir", type=Path, default=Path("artifacts/verification/observed_benchmark"))
    args = parser.parse_args()
    configure_cpu_runtime()
    root = args.project_root.resolve()
    output = (args.output_dir if args.output_dir.is_absolute() else root / args.output_dir).resolve()
    if output.exists() and any(output.iterdir()):
        raise ValueError("Pilot output must be absent or empty; existing evidence is never overwritten")
    output.mkdir(parents=True, exist_ok=True)
    report_path = output / "observed_pilot_report.json"
    report = {
        "contract": "solar-observed-benchmark-pilot.v1",
        "created_at": datetime.now(timezone.utc).isoformat(),
        "status": "started", "data_origin": DATA_ORIGIN,
        "optimization_scope": OPTIMIZATION_SCOPE,
        "smoke": False,
        "full_population_optimization_completed": False,
        "historical_checkpoint_performance_verified": False,
        "limitations": [
            "One plant-year chosen by observation availability, not the complete plant population",
            "Three CNN training epochs; bounded feature, lookback, and hyperparameter search",
            "Original trained checkpoint results are not reproduced by this pilot",
        ],
        "runtime": runtime_versions(),
    }
    write_json(report_path, report)
    try:
        sys.path.insert(0, str(root / "src"))
        from solar_forecast.evaluation.experiment_config import load_experiment_config
        config_path = args.config if args.config.is_absolute() else root / args.config
        base = load_experiment_config(config_path, project_root=root)
        revision = subprocess.run(
            ["git", "-C", str(root), "rev-parse", "HEAD"], capture_output=True, text=True, check=True
        ).stdout.strip()
        report["source_revision"] = revision
        prepared = output / "prepared"
        gold_path = prepare_official_gold(root, prepared)
        import pandas as pd
        cohort, selection = select_observed_plant_year(
            pd.read_csv(gold_path, low_memory=False), split=base.get("split", {}),
            minimum_coverage=float(base.get("minimum_common_coverage", 0.95)),
        )
        dataset = output / "observed_plant_year.csv"
        cohort.to_csv(dataset, index=False, encoding="utf-8-sig")
        provenance = capture_input_provenance(root, prepared, output)
        provenance.update({
            "data_origin": DATA_ORIGIN, "optimization_scope": OPTIMIZATION_SCOPE,
            "source_revision": revision, "cohort": selection,
            "input_dataset": str(dataset), "input_dataset_sha256": sha256_file(dataset),
            "gold_dataset_sha256": sha256_file(gold_path),
        })
        report.update({"status": "prepared", "provenance": provenance})
        write_json(report_path, report)
        values = build_pilot_config(base, dataset, output, root, provenance)
        pilot_config = output / "pilot_experiment.json"
        write_json(pilot_config, values)
        # Validate the exact generated configuration before any training starts.
        load_experiment_config(pilot_config, project_root=root)
        import torch
        torch.set_num_threads(1)
        torch.set_num_interop_threads(1)
        from solar_forecast.jobs.benchmark_job import BenchmarkService
        run_dir = BenchmarkService(project_root=root).run(pilot_config, smoke=False)
        from verify_benchmark_model_artifacts import verify_selected_artifacts
        replay = verify_selected_artifacts(run_dir, dataset)
        write_json(output / "model_artifact_replay.json", replay)
        if replay["status"] != "passed":
            failed_manifest = json.loads((run_dir / "manifest.json").read_text(encoding="utf-8"))
            failed_manifest.update(status="failed", error="Stored model artifact replay failed")
            write_json(run_dir / "manifest.json", failed_manifest)
            raise ValueError("Stored model artifact replay failed; inspect model_artifact_replay.json")
        report["stored_model_artifact_replay"] = replay
        from solar_forecast.reporting.dashboard_builder import DashboardBuilder
        dashboard = DashboardBuilder(root).build()
        payload = json.loads(dashboard.data_path.read_text(encoding="utf-8"))
        benchmark = payload.get("model_benchmark", {})
        tasks = benchmark.get("tasks", [])
        if (
            benchmark.get("status") != "ready"
            or benchmark.get("run_id") != run_dir.name
            or sorted(task.get("horizon_hours") for task in tasks) != values["horizons_hours"]
            or any(task.get("selected_model") not in {"xgboost", "cnn_bilstm", "hybrid"} for task in tasks)
        ):
            raise ValueError("Dashboard does not expose the completed pilot and both selected horizons")
        report.update({
            "status": "completed", "benchmark_run_dir": str(run_dir),
            "dashboard": {
                "data_path": str(dashboard.data_path), "sha256": sha256_file(dashboard.data_path),
                "benchmark_run_id": benchmark["run_id"], "status": "ready",
                "selected_models": {str(task["horizon_hours"]): task["selected_model"] for task in tasks},
            },
        })
        write_json(report_path, report)
        print(json.dumps({
            "status": "completed", "optimization_scope": OPTIMIZATION_SCOPE,
            "cohort": selection, "benchmark_run_dir": str(run_dir), "report": str(report_path),
        }, ensure_ascii=False, indent=2))
    except Exception as exc:
        report.update({"status": "failed", "error_type": type(exc).__name__, "error": str(exc)})
        write_json(report_path, report)
        raise


if __name__ == "__main__":
    main()
