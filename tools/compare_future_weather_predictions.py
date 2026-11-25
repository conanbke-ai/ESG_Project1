"""Compare completed weather-on and weather-off forecasts on identical saved rows."""
from __future__ import annotations

import argparse
import json
from pathlib import Path, PureWindowsPath
import re
import sys

import numpy as np
import pandas as pd

SRC_ROOT = Path(__file__).resolve().parents[1] / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from solar_forecast.evaluation.forecast_samples import forecast_cohort_contract
from solar_forecast.evaluation.model_selection import _metrics
from solar_forecast.evaluation.temporal_split import TemporalSplitConfig, TemporalSplitter
from solar_forecast.features.future_weather import (
    FUTURE_WEATHER_CONTRACT,
    SUPPORTED_FIXED_LEADS,
    future_weather_feature_columns,
)
from solar_forecast.infrastructure.artifact_store import sha256_file
from solar_forecast.jobs.benchmark_job import (
    BENCHMARK_CONTRACT,
    PREDICTION_KEYS,
    align_prediction_frames,
)
from solar_forecast.jobs.contracts import TRAINING_RUN_MANIFEST_CONTRACT


FOLDS = ("validation", "calibration", "test")
FIXED_TREATMENT_WITNESS = {
    "weather_on": (
        "observed_measurements_through_forecast_origin_plus_"
        "archived_fixed_lead_weather_forecast"
    ),
    "weather_off": "observed_measurements_through_forecast_origin",
}


def _read_object(path: Path) -> dict:
    result = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(result, dict):
        raise ValueError(f"{path}: expected a JSON object")
    return result


def _benchmark_provenance(run_dir: Path, dataset: str) -> tuple[Path, dict]:
    for parent in run_dir.parents:
        manifest_path = parent / "manifest.json"
        if not manifest_path.is_file():
            continue
        manifest = _read_object(manifest_path)
        if manifest.get("contract") != BENCHMARK_CONTRACT:
            continue
        provenance = manifest.get("provenance") or {}
        if (
            manifest.get("status") != "completed"
            or manifest.get("execution_mode") != "full"
            or provenance.get("dataset_fingerprint") != dataset
        ):
            raise ValueError(f"{manifest_path}: expected a completed full benchmark for the same dataset")
        revision = str(provenance.get("code_revision", ""))
        source_hash = str(provenance.get("source_sha256", ""))
        if not re.fullmatch(r"[0-9a-f]{40}", revision) or not re.fullmatch(r"[0-9a-f]{64}", source_hash):
            raise ValueError(f"{manifest_path}: benchmark must record code revision and source SHA-256")
        if not isinstance(provenance.get("runtime"), dict) or not provenance["runtime"]:
            raise ValueError(f"{manifest_path}: benchmark runtime versions are missing")
        return manifest_path, {
            key: provenance[key] for key in ("code_revision", "source_sha256", "runtime")
        }
    raise ValueError(f"{run_dir}: preserve the enclosing benchmark manifest with the candidate")


def _settings(config: dict) -> dict:
    # Artifact locations and Optuna study names differ where treatment is forked.
    ignored = {"future_weather", "input_dataset", "output_root", "prediction_root", "benchmark_candidate_id"}
    settings = {key: value for key, value in config.items() if key not in ignored}
    for section, ignored_fields in (
        ("optimizer", {"study_name", "storage_path"}),
        ("checkpoint", {"root"}),
    ):
        if isinstance(settings.get(section), dict):
            settings[section] = {
                key: value for key, value in settings[section].items()
                if key not in ignored_fields
            }
    return settings


def _load_run(run_dir: Path, arm: str) -> dict:
    run_dir = run_dir.resolve()
    manifest_path = run_dir / "manifest.json"
    config_path = run_dir / "resolved_config.json"
    manifest = _read_object(manifest_path)
    config = _read_object(config_path)
    model = manifest.get("model")
    details = manifest.get("details") or {}
    if (
        manifest.get("contract") != TRAINING_RUN_MANIFEST_CONTRACT
        or manifest.get("status") != "completed"
        or model not in ("xgboost", "cnn_bilstm")
        or config.get("model") != model
        or details.get("selection_only")
        or (details.get("run_context") or {}).get("execution_mode") != "full"
    ):
        raise ValueError(f"{arm}: expected a completed full XGBoost or CNN-BiLSTM candidate")

    contract = details.get("evaluation_contract") or {}
    dataset = str(contract.get("dataset_fingerprint", ""))
    horizon = contract.get("horizon_hours")
    if (
        type(horizon) is not int
        or horizon not in SUPPORTED_FIXED_LEADS
        or horizon != config.get("forecast_horizon_hours")
        or not re.fullmatch(r"[0-9a-f]{64}", dataset)
        or contract.get("task") != "historical_forecast"
        or contract.get("prediction_schema") != "solar-forecast-prediction.v2"
        or contract.get("target") != "generation_mwh"
        or config.get("target_column") != "generation_mwh"
        or contract.get("target_unit") != "MWh"
        or contract.get("origin_definition") != "target_timestamp_minus_horizon"
        or contract.get("information_set") != FIXED_TREATMENT_WITNESS[arm]
    ):
        raise ValueError(f"{arm}: historical fixed-lead MWh contract or weather treatment differs")

    features = config.get("feature_columns") or []
    if (
        not isinstance(features, list)
        or not features
        or any(not isinstance(feature, str) or not feature.strip() for feature in features)
        or len(set(features)) != len(features)
    ):
        raise ValueError(f"{arm}: ordered historical feature list is missing or invalid")
    weather = config.get("future_weather") or {}
    if not isinstance(weather, dict):
        raise ValueError(f"{arm}: resolved forecast settings must be a JSON object")
    weather_evidence = details.get("future_weather")
    enabled = weather.get("enabled", False)
    if arm == "weather_on":
        if enabled is not True or not isinstance(weather_evidence, dict):
            raise ValueError(f"{arm}: enabled weather and actual saved feature evidence are required")
        future = list(future_weather_feature_columns(str(weather.get("feature_profile", ""))))
        source_model = weather.get("source_model")
        recorded_models = weather_evidence.get("forecast_models")
        if (
            not isinstance(source_model, str) or not source_model.strip()
            or weather_evidence.get("contract") != FUTURE_WEATHER_CONTRACT
            or contract.get("future_weather_contract") != FUTURE_WEATHER_CONTRACT
            or weather_evidence.get("feature_columns") != future
            or (recorded_models is not None and recorded_models != [source_model])
            or details.get("features") != (features + future if model == "xgboost" else features)
        ):
            raise ValueError(f"{arm}: saved forecast feature evidence is inconsistent with the resolved config")
    elif enabled is not False or weather_evidence or details.get("features") != features:
        raise ValueError(f"{arm}: expected only the resolved historical features")

    optimizer = details.get("optimizer")
    if not isinstance(optimizer, dict) or type(optimizer.get("enabled")) is not bool:
        raise ValueError(f"{arm}: saved optimizer status is missing")
    if optimizer.get("enabled") is True and (
        optimizer.get("selection_data") != "validation_only"
        or optimizer.get("test_usage") != "none"
        or not isinstance(optimizer.get("best_params"), dict)
    ):
        raise ValueError(f"{arm}: optimized settings must have recorded Validation-only selection")
    if optimizer.get("test_usage", "none") != "none":
        raise ValueError(f"{arm}: saved optimizer used Test")
    optimization_settings = {
        key: optimizer.get(key) for key in (
            "enabled", "best_params", "best_iteration", "selection_data",
            "tuning_train_rows", "tuning_validation_rows", "validation_cohort",
        )
    }
    split = details.get("temporal_split") or {}
    split_counts = split.get("counts") or {}
    split_config = TemporalSplitConfig.from_mapping(config)
    if split_config.split_mode != "calendar" or split_config.gap_hours < horizon:
        raise ValueError(f"{arm}: fixed calendar and horizon-safe purge are required")
    splitter = TemporalSplitter(split_config)
    boundaries = splitter.boundaries(pd.Series(dtype="datetime64[ns]"))
    if (
        split.get("boundaries") != boundaries.to_dict()
        or split.get("smoke_split_override")
        or any(type(split_counts.get(fold)) is not int or split_counts[fold] < 1 for fold in ("train", *FOLDS))
    ):
        raise ValueError(f"{arm}: saved temporal boundaries/counts differ from the calendar")
    if model == "cnn_bilstm" and details.get("target_transform", "identity") != config.get("target_transform", "identity"):
        raise ValueError(f"{arm}: CNN target transform differs from resolved config")

    benchmark_path, benchmark = _benchmark_provenance(run_dir, dataset)
    return {
        "run_dir": run_dir,
        "manifest_path": manifest_path,
        "config_path": config_path,
        "benchmark_path": benchmark_path,
        "benchmark": benchmark,
        "model": model,
        "config": config,
        "details": details,
        "contract": contract,
        "optimizer": optimization_settings,
        "split_counts": split_counts,
        "splitter": splitter,
        "boundaries": boundaries,
    }


def _load_fold(run: dict, arm: str, fold: str) -> tuple[pd.DataFrame, Path]:
    saved = run["details"].get(f"{fold}_predictions")
    if not isinstance(saved, str) or not saved.strip():
        raise ValueError(f"{arm}: saved {fold} prediction artifact is missing")
    # Locate an existing candidate's direct or CNN inner-run export under the
    # supplied directory, including Windows artifacts moved to another drive.
    predictions = PureWindowsPath(saved).name
    matches = [path for path in run["run_dir"].rglob(predictions) if path.name == predictions and path.is_file()]
    if len(matches) != 1:
        raise ValueError(f"{arm}: expected exactly one {fold} prediction file named {predictions!r}")
    path = matches[0]
    frame = pd.read_csv(path, dtype={"plant_id": str}, float_precision="round_trip")
    needed = {*PREDICTION_KEYS, "plant", "region", "split", "y_true", "y_pred", "persistence_pred"}
    if missing := needed - set(frame):
        raise ValueError(f"{arm} {fold}: missing prediction columns: {sorted(missing)}")
    if len(frame) != run["split_counts"][fold] or not frame["split"].eq(fold).fillna(False).all():
        raise ValueError(f"{arm} {fold}: row count or split label differs from saved fold")
    for name in ("plant_id", "plant", "region"):
        if frame[name].fillna("").astype(str).str.strip().eq("").any():
            raise ValueError(f"{arm} {fold}: missing {name}")
    for name in ("timestamp", "forecast_origin"):
        times = pd.to_datetime(frame[name], format="mixed", errors="raise")
        if (
            not pd.api.types.is_datetime64_dtype(times)
            or times.isna().any()
            or not times.eq(times.dt.floor("h")).all()
        ):
            raise ValueError(f"{arm} {fold}: {name} must be hourly in the saved local KST calendar")
        frame[name] = times
    horizon = pd.to_numeric(frame["horizon_hours"], errors="raise")
    if not horizon.eq(run["contract"]["horizon_hours"]).all():
        raise ValueError(f"{arm} {fold}: prediction horizon differs from saved horizon")
    frame["horizon_hours"] = horizon.astype("int64")
    labels = run["splitter"].labels(frame["timestamp"], run["boundaries"])
    if not labels.eq(fold).fillna(False).all():
        raise ValueError(f"{arm} {fold}: target falls outside the saved fold calendar or within purge")
    return frame, path


def _score_pair(frame: pd.DataFrame) -> dict:
    scores = {}
    for arm in ("weather_on", "weather_off", "persistence"):
        prediction = frame[arm]
        scores[arm] = {
            **_metrics(frame["y_true"], prediction),
            "bias": float((prediction - frame["y_true"]).mean()),
        }
    scores["delta_weather_on_minus_off"] = {
        name: None if scores["weather_on"][name] is None else (
            scores["weather_on"][name] - scores["weather_off"][name]
        )
        for name in ("mae", "rmse", "r2", "bias")
    }
    return scores


def compare(weather_on: Path, weather_off: Path) -> dict:
    """Verify exact fixed-calendar fold equality, then compute read-only Test scores."""
    runs = {
        "weather_on": _load_run(weather_on, "weather_on"),
        "weather_off": _load_run(weather_off, "weather_off"),
    }
    on, off = runs["weather_on"], runs["weather_off"]
    settings_on, settings_off = _settings(on["config"]), _settings(off["config"])
    differing = sorted(
        key for key in settings_on.keys() | settings_off.keys()
        if key not in settings_on or key not in settings_off
        or json.dumps(settings_on[key], sort_keys=True) != json.dumps(settings_off[key], sort_keys=True)
    )
    if differing:
        raise ValueError(f"Resolved settings differ beyond forecast treatment: {differing}")
    for key in ("benchmark", "model", "optimizer", "split_counts"):
        if on[key] != off[key]:
            raise ValueError(f"Candidate {key} differs between weather arms")
    treatment_fields = {"information_set", "future_weather_contract", "future_weather_spatial_contract"}
    if any(
        on["contract"].get(key) != off["contract"].get(key)
        for key in (on["contract"].keys() | off["contract"].keys()) - treatment_fields
    ):
        raise ValueError("Evaluation contracts differ beyond the forecast information set")

    precision = on["config"].get("numeric_dtype", "float32")
    if precision not in ("float32", "float64"):
        raise ValueError(f"Unsupported common prediction truth precision: {precision}")
    folds = {}
    test = None
    for fold in FOLDS:
        exports = {arm: _load_fold(run, arm, fold) for arm, run in runs.items()}
        aligned, coverage = align_prediction_frames(
            {arm: result[0] for arm, result in exports.items()}, minimum_coverage=1.0,
        )
        treated, control = aligned["weather_on"], aligned["weather_off"]
        for field in ("y_true", "persistence_pred"):
            left = treated[field].to_numpy(dtype=precision)
            right = control[field].to_numpy(dtype=precision)
            if not np.isfinite(left).all() or not np.array_equal(left, right):
                raise ValueError(f"{fold}: {field} differs at the declared common numeric precision")
        folds[fold] = {
            **forecast_cohort_contract(treated),
            "coverage": coverage,
            "predictions_sha256": {arm: sha256_file(result[1]) for arm, result in exports.items()},
        }
        if fold == "test":
            test = treated.assign(
                month=treated["timestamp"].dt.to_period("M").astype(str),
                y_true=treated["y_true"].astype(precision).astype(float),
                weather_on=treated["y_pred"],
                weather_off=control["y_pred"],
                persistence=treated["persistence_pred"].astype(precision).astype(float),
            )
    if not (
        folds["test"]["target_start"] == pd.Timestamp(on["contract"]["test_start"]).isoformat()
        and folds["test"]["target_end"] == pd.Timestamp(on["contract"]["test_end"]).isoformat()
    ):
        raise ValueError("Paired Test period differs from the saved evaluation contract")
    assert test is not None
    return {
        "model": on["model"],
        "horizon_hours": on["contract"]["horizon_hours"],
        "target_unit": "MWh",
        "dataset_fingerprint": on["contract"]["dataset_fingerprint"],
        "source": on["benchmark"],
        "calendar": on["boundaries"].to_dict(),
        "train_rows": on["split_counts"]["train"],
        "train_key_identity_verified": False,
        "historical_operational_availability_verified": False,
        "weather_source_model": on["config"]["future_weather"]["source_model"],
        "weather_feature_profile": on["config"]["future_weather"]["feature_profile"],
        "run_files_sha256": {
            arm: {kind: sha256_file(run[f"{kind}_path"]) for kind in ("manifest", "config", "benchmark")}
            for arm, run in runs.items()
        },
        "folds": folds,
        "test_metrics": {
            "pooled": _score_pair(test),
            "by_plant": {key: _score_pair(group) for key, group in test.groupby("plant_id", sort=True)},
            "by_month": {key: _score_pair(group) for key, group in test.groupby("month", sort=True)},
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--weather-on", required=True, type=Path, help="Weather-on benchmark candidate directory")
    parser.add_argument("--weather-off", required=True, type=Path, help="Weather-off benchmark candidate directory")
    parser.add_argument("--output", type=Path, help="Optional new JSON file for the Test report")
    args = parser.parse_args()
    try:
        result = json.dumps(compare(args.weather_on, args.weather_off), ensure_ascii=False, indent=2, allow_nan=False)
        if args.output is None:
            print(result)
        else:
            with args.output.open("x", encoding="utf-8") as output:
                output.write(result + "\n")
    except (OSError, ValueError, KeyError) as exc:
        parser.exit(1, f"Forecast comparison blocked: {exc}\n")


if __name__ == "__main__":
    main()
