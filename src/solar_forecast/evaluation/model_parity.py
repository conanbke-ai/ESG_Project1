"""Compare aligned prediction artifacts; output parity is not accuracy certification."""
from __future__ import annotations

import csv
from hashlib import sha256
from io import BytesIO, StringIO
from pathlib import Path
import re

import numpy as np
import pandas as pd


REQUIRED_COLUMNS = {"plant_id", "timestamp", "y_true", "y_pred"}
_HOURLY_TIMESTAMP = re.compile(
    r"^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}"
    r"(?::\d{2}(?:\.\d{1,9})?)?(?:Z|[+-]\d{2}:\d{2})?$"
)


def _parse_hourly_timestamps(values: pd.Series, label: str) -> tuple[list, str]:
    """Parse explicit ISO hour keys without guessing a timezone for naive values."""
    timestamps = []
    timezone_modes = set()
    for row_number, value in enumerate(values, start=2):
        if not _HOURLY_TIMESTAMP.fullmatch(value):
            raise ValueError(f"{label}: malformed timestamp at CSV row {row_number}")
        try:
            timestamp = pd.Timestamp(value)
        except (ValueError, OverflowError) as exc:
            raise ValueError(
                f"{label}: malformed timestamp at CSV row {row_number}"
            ) from exc
        if timestamp.minute or timestamp.second or timestamp.microsecond or timestamp.nanosecond:
            raise ValueError(f"{label}: non-hourly timestamp at CSV row {row_number}")
        aware = timestamp.tzinfo is not None
        timezone_modes.add("aware" if aware else "naive")
        timestamps.append(timestamp.tz_convert("UTC") if aware else timestamp)
    if len(timezone_modes) != 1:
        raise ValueError(f"{label}: mixed timezone-aware and naive timestamps")
    return timestamps, timezone_modes.pop()


def _load_prediction_artifact(path: Path, label: str) -> tuple[pd.DataFrame, dict]:
    """Read the exact hashed bytes and reject ambiguous or unusable prediction rows."""
    payload = path.read_bytes()
    try:
        header = next(csv.reader(StringIO(payload.decode("utf-8-sig"))))
    except (UnicodeError, StopIteration, csv.Error) as exc:
        raise ValueError(f"{label}: expected a non-empty UTF-8 prediction CSV") from exc
    if len(header) != len(set(header)):
        raise ValueError(f"{label}: duplicate CSV column names")
    missing = REQUIRED_COLUMNS - set(header)
    if missing:
        raise ValueError(f"{label}: missing required columns: {sorted(missing)}")
    try:
        frame = pd.read_csv(BytesIO(payload), dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except (ValueError, pd.errors.ParserError) as exc:
        raise ValueError(f"{label}: malformed prediction CSV") from exc
    if frame.empty:
        raise ValueError(f"{label}: prediction CSV has no rows")
    if frame["plant_id"].str.strip().eq("").any():
        raise ValueError(f"{label}: plant_id contains an empty identifier")
    if "region" in frame and frame["region"].str.strip().eq("").any():
        raise ValueError(f"{label}: region contains an empty label")
    timestamps, timezone_mode = _parse_hourly_timestamps(frame["timestamp"], label)
    frame["timestamp"] = timestamps
    for column in ("y_true", "y_pred"):
        try:
            frame[column] = pd.to_numeric(frame[column], errors="raise").astype(float)
        except (ValueError, TypeError, OverflowError) as exc:
            raise ValueError(f"{label}: {column} must contain finite numeric values") from exc
        if not np.isfinite(frame[column].to_numpy()).all():
            raise ValueError(f"{label}: {column} must contain finite numeric values")
    if frame.duplicated(["plant_id", "timestamp"]).any():
        raise ValueError(f"{label}: duplicate plant_id/timestamp keys")
    return frame.set_index(["plant_id", "timestamp"]).sort_index(), {
        "path": str(path),
        "sha256": sha256(payload).hexdigest(),
        "rows": len(frame),
        "timezone_mode": timezone_mode,
    }


def _calculate_error_metrics(actual: np.ndarray, predicted: np.ndarray) -> dict:
    """Compute pooled row metrics, preserving undefined R2 as JSON null."""
    with np.errstate(over="raise", invalid="raise", divide="raise"):
        try:
            errors = actual - predicted
            sum_squared_error = np.square(errors).sum()
            squared_target_variation = np.square(actual - actual.mean()).sum()
            return {
                "n_samples": len(actual),
                "mae": float(np.abs(errors).mean()),
                "rmse": float(np.sqrt(sum_squared_error / len(actual))),
                "r2": (
                    float(1.0 - sum_squared_error / squared_target_variation)
                    if len(actual) > 1 and squared_target_variation > 0 else None
                ),
            }
        except FloatingPointError as exc:
            raise ValueError("Prediction values overflow metric computation") from exc


def _compare_aligned_rows(frame: pd.DataFrame, *, atol: float, rtol: float) -> dict:
    """Measure candidate differences using baseline predictions as the reference."""
    actual = frame["y_true"].to_numpy(float)
    baseline = frame["baseline_prediction"].to_numpy(float)
    candidate = frame["candidate_prediction"].to_numpy(float)
    baseline_metrics = _calculate_error_metrics(actual, baseline)
    candidate_metrics = _calculate_error_metrics(actual, candidate)
    with np.errstate(over="raise", invalid="raise"):
        try:
            differences = np.abs(candidate - baseline)
            matching = np.isclose(candidate, baseline, atol=atol, rtol=rtol)
            mean_absolute_difference = float(differences.mean())
        except FloatingPointError as exc:
            raise ValueError("Prediction values or tolerances overflow comparison") from exc
    return {
        "rows": len(frame),
        "status": "passed" if matching.all() else "failed",
        "prediction_difference": {
            "max_absolute": float(differences.max()),
            "mean_absolute": mean_absolute_difference,
            "rows_outside_tolerance": int((~matching).sum()),
        },
        "baseline": baseline_metrics,
        "candidate": candidate_metrics,
        "delta": {
            metric: (
                candidate_metrics[metric] - baseline_metrics[metric]
                if baseline_metrics[metric] is not None and candidate_metrics[metric] is not None
                else None
            ) for metric in ("mae", "rmse", "r2")
        },
    }


def compare_prediction_files(
    baseline: Path,
    candidate: Path,
    *,
    atol: float = 1e-6,
    rtol: float = 1e-6,
) -> dict:
    """Compare identical observed plant-hours; invalid input raises ValueError.

    A passed result establishes numerical output parity for these two artifacts
    only. Dataset provenance, temporal leakage, training reproducibility, forecast
    horizon, and production performance need independent run-level validation.
    Metric deltas are candidate minus baseline; allclose uses baseline as reference.
    """
    for name, value in (("atol", atol), ("rtol", rtol)):
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not np.isfinite(value) or value < 0:
            raise ValueError(f"{name} must be a finite nonnegative number")
    baseline_frame, baseline_info = _load_prediction_artifact(Path(baseline), "baseline")
    candidate_frame, candidate_info = _load_prediction_artifact(Path(candidate), "candidate")
    if baseline_info["timezone_mode"] != candidate_info["timezone_mode"]:
        raise ValueError("Baseline and candidate timestamp timezone modes differ")
    if not baseline_frame.index.equals(candidate_frame.index):
        missing = len(baseline_frame.index.difference(candidate_frame.index))
        extra = len(candidate_frame.index.difference(baseline_frame.index))
        raise ValueError(
            "Baseline and candidate plant_id/timestamp key sets differ: "
            f"missing candidate keys={missing}, extra candidate keys={extra}"
        )
    if not np.array_equal(baseline_frame["y_true"], candidate_frame["y_true"]):
        raise ValueError("Baseline and candidate observed y_true values differ")
    region_available = "region" in baseline_frame
    if region_available != ("region" in candidate_frame):
        raise ValueError("Baseline and candidate region column presence differs")
    if region_available and not baseline_frame["region"].equals(candidate_frame["region"]):
        raise ValueError("Baseline and candidate region labels disagree")
    if ("split" in baseline_frame) != ("split" in candidate_frame):
        raise ValueError("Baseline and candidate split column presence differs")
    if "split" in baseline_frame and not baseline_frame["split"].equals(candidate_frame["split"]):
        raise ValueError("Baseline and candidate evaluation split labels disagree")
    aligned = baseline_frame[["y_true"]].copy()
    aligned["baseline_prediction"] = baseline_frame["y_pred"]
    aligned["candidate_prediction"] = candidate_frame["y_pred"]
    if region_available:
        aligned["region"] = baseline_frame["region"]
    overall = _compare_aligned_rows(aligned, atol=atol, rtol=rtol)
    return {
        "contract": "solar-prediction-parity.v1",
        "status": overall["status"],
        "scope": "prediction_output_parity_only",
        "production_accuracy_verified": False,
        "inputs": {"baseline": baseline_info, "candidate": candidate_info},
        "tolerances": {
            "atol": float(atol), "rtol": float(rtol),
            "reference": "baseline",
            "rule": "abs(candidate - baseline) <= atol + rtol * abs(baseline)",
        },
        "metric_delta_direction": "candidate_minus_baseline",
        "overall": overall,
        "per_plant": [
            {"plant_id": plant_id, **_compare_aligned_rows(group, atol=atol, rtol=rtol)}
            for plant_id, group in aligned.groupby(level="plant_id", sort=True)
        ],
        "per_region": {
            "status": "available" if region_available else "unavailable",
            "reason": None if region_available else "Prediction CSVs contain no region labels",
            "groups": [
                {"region": region, **_compare_aligned_rows(group, atol=atol, rtol=rtol)}
                for region, group in aligned.groupby("region", sort=True)
            ] if region_available else [],
        },
        "contextual_metrics": {
            "daylight": {"status": "unavailable", "reason": "No validated daylight classification contract supplied"},
            "season": {"status": "unavailable", "reason": "No season grouping policy supplied"},
            "weather": {"status": "unavailable", "reason": "No validated weather regime or missing-weather context supplied"},
        },
    }
