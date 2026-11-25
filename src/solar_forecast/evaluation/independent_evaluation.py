"""Audit an admitted holdout and the frozen comparison protocol without fitting models."""
from __future__ import annotations

from collections import defaultdict
from hashlib import sha256
import json
import math
from pathlib import Path
import re

import pandas as pd

from solar_forecast.config_loader import PROJECT_ROOT
from solar_forecast.evaluation.experiment_config import load_experiment_config
from solar_forecast.evaluation.model_selection import _metrics, _validate_frame
from solar_forecast.evaluation.temporal_split import CALENDAR_FIELDS, TemporalSplitConfig
from solar_forecast.infrastructure.artifact_store import sha256_file, write_json_atomic


PREDICTION_COLUMNS = {
    "xgboost": "xgb_pred", "cnn_bilstm": "cnn_pred",
    "hybrid": "hybrid_pred", "persistence": "persistence_pred",
}
PREDICTION_KEY = ["timestamp", "plant_id", "forecast_origin", "horizon_hours"]
HORIZONS = (1, 24, 72)
PURGE_HOURS = 168
HOUR = pd.Timedelta(hours=1)
YEAR_START = pd.Timestamp("2026-01-01T00:00:00")
YEAR_END = pd.Timestamp("2026-12-31T23:00:00")
DEVELOPMENT_STATUS = "historical_development_evaluation_not_unseen_confirmation"
ADMISSION_CONTRACT = "solar-gold-candidate-admission.v1"
EXPOSURE_CONTRACT = "solar-independent-evaluation-exposure-review.v1"


def _json(path: Path) -> dict:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"Expected a JSON object in {path}")
    return value


def _window(record: dict) -> tuple[pd.Timestamp, pd.Timestamp]:
    endpoints = []
    for name in ("start", "end"):
        raw_value = record.get(name)
        if not isinstance(raw_value, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:00(?::00)?", raw_value):
            raise ValueError(f"{name} must be an ISO KST-local hourly timestamp")
        parsed = pd.Timestamp(raw_value)
        if pd.isna(parsed) or parsed.tzinfo is not None or parsed != parsed.floor("h"):
            raise ValueError(f"{name} must be an ISO KST-local hourly timestamp")
        endpoints.append(parsed)
    start, end = endpoints
    if start > end:
        raise ValueError("Window start must not follow window end")
    return start, end


def _check_hash(path: Path, expected: str) -> str:
    digest = sha256_file(path)
    if digest != expected:
        raise ValueError(f"SHA-256 mismatch for {path}")
    return digest


def _saved_gate_integrity(path: Path, expected: str) -> dict:
    data = path.read_bytes()
    observed = sha256(data).hexdigest()
    if observed == expected:
        return {"status": "byte_exact", "checkout_sha256": observed, "recorded_sha256": expected}
    # Git can store a Windows CSV as LF. Describe what happened to the archived
    # evidence without changing it or treating the checkout hash as an exact match.
    if b"\r" not in data and sha256(data.replace(b"\n", b"\r\n")).hexdigest() == expected:
        return {"status": "crlf_reconstruction_matches", "checkout_sha256": observed, "recorded_sha256": expected}
    raise ValueError(f"SHA-256 mismatch for stored gate profiles {path}")


def _split(config: dict) -> dict:
    settings = config["split"]
    split = TemporalSplitConfig.from_mapping(settings)
    if split.split_mode != "calendar" or split.gap_hours != PURGE_HOURS or config["selection_gap_hours"] != PURGE_HOURS:
        raise ValueError("Independent evaluation requires fixed Train/Validation/Calibration/Test and both 168h purges")
    return {**{field: settings[field] for field in CALENDAR_FIELDS}, "purge_gap_hours": split.gap_hours}


def _saved_evaluation(directory: Path, current_split: dict) -> dict:
    run_path, experiment_path = directory / "manifest.json", directory / "experiment.json"
    manifest, experiment = _json(run_path), _json(experiment_path)
    provenance = manifest["provenance"]
    if (manifest["contract"] != "solar-optimized-benchmark.v1" or manifest["status"] != "completed"
            or provenance["data_provenance"]["evaluation_status"] != DEVELOPMENT_STATUS):
        raise ValueError("Stored benchmark must be a completed development-exposed comparison")
    split = _split(experiment)
    if split != current_split:
        raise ValueError("Stored and current calendar split boundaries differ")
    historical = []
    for horizon in HORIZONS:
        selection_path = directory / f"horizon_{horizon}h" / "selection.json"
        selection = _json(selection_path)
        contract = selection["evaluation_contract"]
        if (selection["status"] != "completed" or contract["horizon_hours"] != horizon
                or contract["target_unit"] != "MWh" or contract["target"] != "generation_mwh"
                or contract["dataset_fingerprint"] != provenance["dataset_fingerprint"]
                or contract["prediction_key"] != PREDICTION_KEY
                or selection["selection_rule"]["test_used_for_selection"] is not False
                or selection["purge"]["selection_gap_hours"] != PURGE_HOURS):
            raise ValueError(f"Stored selection contract differs for horizon {horizon}")
        files = {"test_predictions": "test_predictions.csv.gz", "gate_profiles": "gate_profiles.csv"}
        if selection["files"] != files:
            raise ValueError(f"Unexpected saved filenames at horizon {horizon}")
        test_path = selection_path.parent / files["test_predictions"]
        test_hash = _check_hash(test_path, selection["files_sha256"]["test_predictions"])
        gate_integrity = _saved_gate_integrity(selection_path.parent / files["gate_profiles"],
                                               selection["files_sha256"]["gate_profiles"])
        columns = {*PREDICTION_KEY, *PREDICTION_COLUMNS.values(), "y_true", "plant", "region", "split"}
        frame = pd.read_csv(test_path, usecols=columns, float_precision="round_trip")
        frame = _validate_frame(frame, "saved 2025 development Test", horizon)
        period = selection["periods"]["test"]
        start, end = _window(period)
        if (not frame["split"].eq("test").all() or frame.timestamp.min() != start or frame.timestamp.max() != end
                or frame.timestamp.dt.tz is not None or len(frame) != period["rows"]
                or frame.plant_id.nunique() != period["plants"] or frame.timestamp.nunique() != period["unique_timestamps"]
                or start <= pd.Timestamp(split["calibration_end"]) + PURGE_HOURS * HOUR
                or end > pd.Timestamp(split["test_end"])):
            raise ValueError(f"Stored development Test coverage differs for horizon {horizon}")
        saved = {}
        for model, column in PREDICTION_COLUMNS.items():
            recomputed = _metrics(frame.y_true, frame[column])
            pooled = selection["test_metrics"][model]["pooled"]
            for key, actual in recomputed.items():
                recorded = pooled[key]
                equal = actual == recorded if actual is None or recorded is None else math.isclose(
                    actual, recorded, rel_tol=1e-12, abs_tol=1e-12)
                if not equal:
                    raise ValueError(f"Stored development {key} differs for {model} at horizon {horizon}")
            saved[model] = pooled
        historical.append({"horizon_hours": horizon, "status": DEVELOPMENT_STATUS, "period": period,
                           "information_set": contract["information_set"],
                           "observation_availability_assumption": contract["observation_availability_assumption"],
                           "pooled_saved_metrics_verified": saved, "selection_sha256": sha256_file(selection_path),
                           "test_predictions_sha256": test_hash, "gate_profiles_integrity": gate_integrity})
    return {"run_id": directory.name, "manifest_sha256": sha256_file(run_path),
            "experiment_sha256": sha256_file(experiment_path), "dataset_fingerprint": provenance["dataset_fingerprint"],
            "source_code_revision": provenance["code_revision"], "horizons": historical}


def _admitted_periods(candidate: dict) -> list[tuple[str, pd.Timestamp, pd.Timestamp]]:
    periods, by_plant = [], defaultdict(list)
    rows = candidate["admitted_2026_rows"]
    if type(rows) is not int or rows < 0:
        raise ValueError("admitted_2026_rows must be a nonnegative integer")
    counted = 0
    for item in candidate["eligible_periods"]:
        start, end = _window(item)
        plant_id = item["plant_id"]
        count = int((end - start) / HOUR) + 1
        if (not isinstance(plant_id, str) or not plant_id.strip() or type(item["rows"]) is not int
                or item["rows"] != count or start < YEAR_START or end > YEAR_END):
            raise ValueError("Admitted periods must enumerate verified contiguous 2026 plant hours")
        periods.append((plant_id, start, end))
        by_plant[plant_id].append((start, end))
        counted += count
    if counted != rows or (rows and (candidate["status"] != "admitted" or candidate.get("period_end_inclusive") is not True)):
        raise ValueError("Admitted period rows do not agree with the candidate manifest")
    for windows in by_plant.values():
        ordered = sorted(windows)
        if any(next_start <= end for (_, end), (next_start, _) in zip(ordered, ordered[1:])):
            raise ValueError("Admitted periods for a plant overlap")
    return sorted(periods)


def _unseen_periods(periods: list, exposure_path: Path, data_version: str, admission_digest: str,
                    development_end: pd.Timestamp) -> tuple[list[dict], dict]:
    exposure = _json(exposure_path)
    if (exposure.get("contract") != EXPOSURE_CONTRACT or exposure.get("status") != "complete"
            or exposure.get("data_version") != data_version or exposure.get("admission_manifest_sha256") != admission_digest):
        raise ValueError("Exposure review must be complete and tied to this candidate manifest")
    window_start, window_end = _window(exposure["reviewed_period"])
    if any(start < window_start or end > window_end for _, start, end in periods):
        raise ValueError("Exposure review must cover all admitted periods")
    sources = exposure["sources"]
    if not sources:
        raise ValueError("Exposure review requires hashed development-use evidence")
    for source in sources:
        _check_hash(exposure_path.parent / source["path"], source["sha256"])
    for item in exposure["exposed_periods"]:
        start, end = _window(item)
        if start < window_start or end > window_end:
            raise ValueError("Exposed periods must lie within the reviewed period")
        development_end = max(development_end, end)
    earliest = development_end + (PURGE_HOURS + 1) * HOUR
    eligible = []
    for plant_id, start, end in periods:
        start = max(start, earliest)
        if start <= end:
            eligible.append({"plant_id": plant_id, "start": start.isoformat(), "end": end.isoformat(),
                             "candidate_target_rows": int((end - start) / HOUR) + 1})
    return eligible, {"sha256": sha256_file(exposure_path), "sources_verified": len(sources),
                      "latest_development_target": development_end.isoformat(),
                      "first_target_after_purge": earliest.isoformat()}


def run_independent_evaluation_audit(*, config_path: Path, development_run: Path, admission_path: Path,
                                     exposure_path: Path | None = None, output_path: Path | None = None,
                                     project_root: Path = PROJECT_ROOT) -> dict:
    """Review metadata and saved development outputs without opening holdout labels or fitting anything."""
    config_path = project_root / config_path
    inputs = [config_path, development_run, admission_path]
    if exposure_path is not None:
        inputs.append(exposure_path)
    if output_path is not None:
        output = output_path.resolve()
        if output.exists() or any(path.resolve() == output or path.resolve() in output.parents for path in inputs):
            raise FileExistsError("Use a new output file outside all input artifacts")
    config = load_experiment_config(config_path, project_root=project_root)
    split = _split(config)
    if sorted(config["horizons_hours"]) != list(HORIZONS):
        raise ValueError("Independent comparison requires exactly 1h, 24h and 72h")
    development = _saved_evaluation(development_run, split)
    admission = _json(admission_path)
    digest = sha256_file(admission_path)
    if (admission.get("contract") != ADMISSION_CONTRACT or admission.get("horizons_hours") != list(HORIZONS)
            or admission.get("time_zone") != "Asia/Seoul"):
        raise ValueError("Unexpected P1 admission, horizon or timezone contract")
    version = admission["data_version"]
    candidate = admission["candidate"]
    periods = _admitted_periods(candidate)
    verified, exposure = [], None
    if not periods:
        status = "blocked_no_admitted_2026_targets"
    else:
        _check_hash(project_root / candidate["dataset_path"], candidate["dataset_sha256"])
        if exposure_path is None:
            status = "blocked_exposure_review_pending"
        else:
            verified, exposure = _unseen_periods(periods, exposure_path, version, digest, pd.Timestamp(split["test_end"]))
            status = "candidate_periods_verified" if verified else "blocked_no_period_after_development_purge"

    future_weather = config.get("future_weather") or {}
    tasks = []
    for horizon in HORIZONS:
        weather = (future_weather.get("by_horizon") or {}).get(str(horizon))
        tasks.append({"horizon_hours": horizon, "models": list(PREDICTION_COLUMNS), "prediction_key": PREDICTION_KEY,
                      "forecast_weather_enabled": horizon in future_weather.get("enabled_horizons", []),
                      "forecast_weather_spec": weather, "common_coverage_minimum": config["minimum_common_coverage"],
                      "configured_forecast_archive_found": (project_root / weather["source"]).is_file() if weather else None,
                      "forecast_weather_coverage_minimum": future_weather.get("minimum_coverage") if weather else None,
                      "preflight_status": "pending" if verified else "blocked_by_holdout_gate",
                      "scoring_status": "pending"})
    report = {
        "contract": "solar-independent-evaluation-readiness.v1", "status": status,
        "development_evaluation": development,
        "holdout": {"candidate_data_version": version, "candidate_manifest_sha256": digest,
                    "p1_status": candidate["status"], "p1_admitted_rows": candidate["admitted_2026_rows"],
                    "p1_eligible_period_count": len(periods), "exposure_review": exposure,
                    "verified_candidate_periods": verified,
                    "candidate_target_rows": sum(period["candidate_target_rows"] for period in verified),
                    "prediction_cohort_rows": None},
        "comparison_protocol": {"experiment_config_sha256": sha256_file(config_path), "frozen_legacy_split": split,
                                "hybrid_selection_purge_hours": PURGE_HOURS,
                                "roles": {"train": "fit_weights_and_preprocessing_scalers_only",
                                          "validation": "choose_base_models_and_hyperparameters",
                                          "calibration": "fit_hybrid_early_then_select_late_after_purge",
                                          "test": "score_frozen_models_once_without_selection"},
                                "preprocessing": {
                                    "cnn_bilstm": "reuse_saved_train_only_imputation_scaler_and_feature_order",
                                    "xgboost": "reuse_saved_feature_order_and_native_missing_value_policy",
                                    "refit_on_holdout": False,
                                    "scoring_units": "convert_each_prediction_back_to_MWh_before_alignment",
                                },
                                "target": "one_hour_generation_mwh_at_target_timestamp",
                                "metrics": ["mae_mwh", "rmse_mwh", "r2"],
                                "metric_scopes": ["pooled_plant_hour", "per_plant", "regional_hourly_sum", "cohort_hourly_sum"],
                                "base_families_share_information_cutoffs": True,
                                "frozen_origin_policy": {
                                    "persistence": "measured_generation_at_exact_target_minus_horizon",
                                    "cnn_window_end": "includes_observation_at_exact_forecast_origin",
                                    "if_origin_measurement_arrives_late": "exclude_key_from_all_four_families",
                                    "reason": "interval-start timestamp does not establish publication by the origin",
                                    "switch_to_latest_available_requires_new_input_contract": True,
                                },
                                "required_preflight": ["freeze_and_hash_training_artifacts_feature_order_and_calibration_gate",
                                                       "freeze_prediction_keys_and_same_plant_hour_truth_before_opening_labels",
                                                       "require_every_input_observation_available_by_origin",
                                                       "require_actual_forecast_issue_at_or_before_receive_at_or_before_origin_where_enabled",
                                                       "require_exact_continuous_cnn_lookback_and_available_origin_persistence",
                                                       "report_common_key_coverage_and_exclusions_per_plant_and_horizon",
                                                       "report_unseen_plants_as_separate_cold_start_cohort"],
                                "tasks": tasks},
        "holdout_labels_read": False, "new_training_performed": False, "independent_scoring_performed": False,
    }
    if output_path is not None:
        write_json_atomic(output_path, report)
    return report
