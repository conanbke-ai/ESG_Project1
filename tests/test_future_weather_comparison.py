"""Regression coverage for the read-only, fixed-cohort forecast comparison."""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest

from solar_forecast.evaluation.forecast_samples import forecast_evaluation_contract
from solar_forecast.evaluation.temporal_split import TemporalSplitConfig, TemporalSplitter
from solar_forecast.features.future_weather import FUTURE_WEATHER_CONTRACT, future_weather_feature_columns
from solar_forecast.jobs.benchmark_job import BENCHMARK_CONTRACT
from solar_forecast.jobs.contracts import TRAINING_RUN_MANIFEST_CONTRACT
from tools.compare_future_weather_predictions import compare


def _mutate_json(path: Path, change) -> None:
    data = json.loads(path.read_text(encoding="utf-8"))
    change(data)
    path.write_text(json.dumps(data), encoding="utf-8")


def _save_candidate(tmp_path: Path, enabled: bool) -> Path:
    benchmark = tmp_path / ("weather-on" if enabled else "weather-off")
    run = benchmark / "candidates" / "candidate"
    run.mkdir(parents=True)
    dataset_fingerprint = "a" * 64
    parent_manifest = {
        "contract": BENCHMARK_CONTRACT,
        "status": "completed",
        "execution_mode": "full",
        "provenance": {
            "dataset_fingerprint": dataset_fingerprint,
            "code_revision": "b" * 40,
            "source_sha256": "c" * 64,
            "runtime": {"xgboost": "3.1.0"},
        },
    }
    (benchmark / "manifest.json").write_text(json.dumps(parent_manifest), encoding="utf-8")
    features = ["capacity_mw"]
    profile = "aligned_core"
    config = {
        "model": "xgboost",
        "target_column": "generation_mwh",
        "forecast_horizon_hours": 24,
        "feature_columns": features,
        "seed": 42,
        "train_end": "2024-06-30T23:00",
        "validation_end": "2024-09-30T23:00",
        "calibration_end": "2024-12-31T23:00",
        "test_end": "2025-12-31T23:00",
        "purge_gap_hours": 168,
        "output_root": str(benchmark),
        "optimizer": {"enabled": True, "max_trials": 5, "study_name": benchmark.name},
        "future_weather": {"enabled": enabled},
    }
    details = {
        "run_context": {"execution_mode": "full"},
        "features": list(features),
        "evaluation_contract": {
            **forecast_evaluation_contract("historical_forecast", 24, legacy_task=""),
            "dataset_fingerprint": dataset_fingerprint,
            "target": "generation_mwh",
            "target_unit": "MWh",
            "test_start": "2025-01-08T00:00:00",
            "test_end": "2025-02-01T00:00:00",
        },
        "temporal_split": {
            "boundaries": TemporalSplitter(TemporalSplitConfig.from_mapping(config)).boundaries(
                pd.Series(dtype="datetime64[ns]")
            ).to_dict(),
            "counts": {"train": 15, "validation": 2, "calibration": 2, "test": 2},
        },
        "optimizer": {
            "enabled": True,
            "selection_data": "validation_only",
            "test_usage": "none",
            "best_params": {"max_depth": 4},
        },
    }
    if enabled:
        future = list(future_weather_feature_columns(profile))
        config["future_weather"].update(source_model="jma_msm", feature_profile=profile)
        details["features"].extend(future)
        details["evaluation_contract"].update({
            "information_set": (
                "observed_measurements_through_forecast_origin_plus_"
                "archived_fixed_lead_weather_forecast"
            ),
            "future_weather_contract": FUTURE_WEATHER_CONTRACT,
        })
        details["future_weather"] = {
            "contract": FUTURE_WEATHER_CONTRACT,
            "forecast_models": ["jma_msm"],
            "feature_columns": future,
        }
    (run / "resolved_config.json").write_text(json.dumps(config), encoding="utf-8")

    for fold, start, end in (
        ("validation", "2024-07-08", "2024-08-01"),
        ("calibration", "2024-10-08", "2024-11-01"),
        ("test", "2025-01-08", "2025-02-01"),
    ):
        timestamps = pd.to_datetime([start, end])
        rows = pd.DataFrame({
            "split": fold,
            "plant_id": ["plant-1", "plant-2"],
            "plant": ["First", "Second"],
            "region": ["South", "North"],
            "timestamp": timestamps,
            "forecast_origin": timestamps - pd.Timedelta(hours=24),
            "horizon_hours": 24,
            "y_true": [1.0, 3.0],
            "y_pred": [1.0, 3.5] if enabled else [2.0, 2.0],
            "persistence_pred": [0.0, 2.0],
        })
        if not enabled:
            rows = rows.iloc[::-1]
        filename = f"{fold}_predictions.csv"
        rows.to_csv(run / filename, index=False)
        details[f"{fold}_predictions"] = rf"C:\original\candidate\{filename}"
    manifest = {
        "contract": TRAINING_RUN_MANIFEST_CONTRACT,
        "status": "completed",
        "model": "xgboost",
        "details": details,
    }
    (run / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
    return run


@pytest.fixture
def paired_runs(tmp_path):
    return _save_candidate(tmp_path, True), _save_candidate(tmp_path, False)


def test_reports_exact_folds_and_paired_metrics(paired_runs):
    report = compare(*paired_runs)

    assert set(report["folds"]) == {"validation", "calibration", "test"}
    assert all(fold["coverage"]["common_fraction"] == 1 for fold in report["folds"].values())
    test = report["test_metrics"]
    assert test["pooled"]["weather_on"] == {
        "n_samples": 2, "mae": 0.25, "rmse": pytest.approx((0.25 / 2) ** 0.5), "r2": 0.875, "bias": 0.25,
    }
    assert test["pooled"]["persistence"]["mae"] == 1.0
    assert test["pooled"]["delta_weather_on_minus_off"]["mae"] == -0.75
    assert test["by_plant"]["plant-2"]["weather_on"]["r2"] is None
    assert set(test["by_month"]) == {"2025-01", "2025-02"}
    assert report["train_key_identity_verified"] is False


@pytest.mark.parametrize(("fold", "field", "value", "error"), [
    ("validation", "plant_id", "unknown", "Common prediction coverage"),
    ("calibration", "y_true", 42.0, "inconsistent y_true"),
    ("test", "persistence_pred", 14, "inconsistent persistence_pred"),
    ("test", "horizon_hours", 24.5, "prediction horizon"),
    ("test", "split", "calibration", "split label"),
    ("test", "timestamp", "2024-12-31", "outside the saved fold calendar"),
])
def test_rejects_mismatched_cohort_or_actuals(paired_runs, fold, field, value, error):
    _, control = paired_runs
    path = control / f"{fold}_predictions.csv"
    frame = pd.read_csv(path)
    if field == "horizon_hours":
        frame[field] = frame[field].astype(float)
    frame.loc[0, field] = value
    frame.to_csv(path, index=False)

    with pytest.raises(ValueError, match=error):
        compare(*paired_runs)


def test_rejects_different_hyperparameters(paired_runs):
    _, control = paired_runs
    path = control / "manifest.json"
    _mutate_json(path, lambda data: data["details"]["optimizer"]["best_params"].update(max_depth=8))

    with pytest.raises(ValueError, match="optimizer differs"):
        compare(*paired_runs)


def test_rejects_different_source_revision(paired_runs):
    _, control = paired_runs
    _mutate_json(control.parents[1] / "manifest.json", lambda data: data["provenance"].update(code_revision="d" * 40))

    with pytest.raises(ValueError, match="benchmark differs"):
        compare(*paired_runs)


def test_rejects_claimed_weather_without_actual_features(paired_runs):
    treated, _ = paired_runs
    _mutate_json(treated / "manifest.json", lambda data: data["details"].pop("future_weather"))

    with pytest.raises(ValueError, match="actual saved feature evidence"):
        compare(*paired_runs)
