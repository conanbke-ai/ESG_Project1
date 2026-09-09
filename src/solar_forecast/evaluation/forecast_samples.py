"""실측 이력의 예측 기준시각·목표시각 정렬과 연속 시간창 검증."""
from __future__ import annotations

from numbers import Integral
from typing import Sequence

import numpy as np
import pandas as pd


HISTORICAL_FORECAST_TASK = "historical_forecast"
FORECAST_CONTEXT_COLUMNS = ("forecast_origin", "horizon_hours", "persistence_pred")


def validate_forecast_horizon(value: object) -> int:
    """Reject ambiguous horizons instead of silently truncating fractional hours."""

    if isinstance(value, bool) or not isinstance(value, Integral) or value < 1:
        raise ValueError("forecast_horizon_hours must be a positive integer")
    return int(value)


def validate_observation_frame(
    frame: pd.DataFrame, *, entity_column: str, timestamp_column: str
) -> pd.DataFrame:
    """Require a unique observed plant-hour before constructing forecast targets."""

    missing = {entity_column, timestamp_column}.difference(frame.columns)
    if missing:
        raise ValueError(f"Forecast identity columns are missing: {sorted(missing)}")
    prepared = frame.copy()
    prepared[timestamp_column] = pd.to_datetime(prepared[timestamp_column], errors="raise")
    if prepared[[entity_column, timestamp_column]].isna().any().any():
        raise ValueError("Forecast observations require non-null plant and timestamp")
    timestamps = prepared[timestamp_column]
    if not timestamps.eq(timestamps.dt.floor("h")).all():
        raise ValueError("Forecast observations must use hourly timestamps")
    if prepared.duplicated([entity_column, timestamp_column]).any():
        raise ValueError("Duplicate plant-hour observations cannot define a forecast origin")
    return prepared.sort_values([entity_column, timestamp_column], kind="stable")


def build_forecast_samples(
    frame: pd.DataFrame,
    feature_columns: Sequence[str],
    target_column: str,
    horizon_hours: int,
    *,
    entity_column: str = "plant_id",
    timestamp_column: str = "timestamp",
) -> pd.DataFrame:
    """Pair y(t) with features observed at exactly t-h, independently per plant.

    No interpolation, forward filling, row-count shift, or future ASOS value is
    used. Calendar-derived features in the input are also taken at the origin.
    The persistence benchmark is the measured generation at that same origin.
    """

    horizon = validate_forecast_horizon(horizon_hours)
    features = list(feature_columns)
    if target_column in features:
        raise ValueError("The future target cannot also be an unshifted feature column")
    prepared = validate_observation_frame(
        frame, entity_column=entity_column, timestamp_column=timestamp_column
    )
    metadata_columns = [column for column in prepared if column not in features]
    target_rows = prepared[metadata_columns].copy()
    target_rows["forecast_origin"] = target_rows[timestamp_column] - pd.Timedelta(hours=horizon)
    origin_rows = prepared[[entity_column, timestamp_column, *features, target_column]].rename(
        columns={timestamp_column: "forecast_origin", target_column: "persistence_pred"}
    )
    samples = target_rows.merge(
        origin_rows,
        on=[entity_column, "forecast_origin"],
        how="inner",
        validate="one_to_one",
        sort=False,
    )
    samples = samples.loc[
        np.isfinite(pd.to_numeric(samples[target_column], errors="coerce"))
        & np.isfinite(pd.to_numeric(samples["persistence_pred"], errors="coerce"))
    ].copy()
    samples["horizon_hours"] = horizon
    return samples.sort_values([timestamp_column, entity_column], kind="stable").reset_index(drop=True)


def forecast_window_positions(
    timestamps: Sequence[object], *, horizon_hours: int, sequence_length: int
) -> tuple[np.ndarray, np.ndarray]:
    """Return target and origin positions whose input window is truly hourly.

    The last input is the observation at origin, included in the window. Gaps
    anywhere inside that window invalidate it; missing future timestamps are
    never invented to provide a target.
    """

    horizon = validate_forecast_horizon(horizon_hours)
    if isinstance(sequence_length, bool) or not isinstance(sequence_length, Integral) or sequence_length < 1:
        raise ValueError("sequence_length must be a positive integer")
    times = pd.DatetimeIndex(pd.to_datetime(timestamps, errors="raise"))
    if times.hasnans or not times.is_unique or not times.is_monotonic_increasing:
        raise ValueError("Forecast window timestamps must be valid, unique, and sorted")
    if not times.equals(times.floor("h")):
        raise ValueError("Forecast windows require hourly timestamps")
    origins = times.get_indexer(times - pd.Timedelta(hours=horizon))
    targets = np.arange(len(times), dtype=np.int64)
    enough_history = origins >= sequence_length - 1
    targets, origins = targets[enough_history], origins[enough_history]
    # A cumulative count detects every discontinuity in O(rows), without
    # materializing a matrix of overlapping windows.
    gaps = np.zeros(len(times), dtype=np.int64)
    if len(times) > 1:
        gaps[1:] = np.asarray(times[1:] - times[:-1]) != np.timedelta64(1, "h")
    gaps = np.cumsum(gaps)
    starts = origins - sequence_length + 1
    continuous = gaps[origins] == gaps[starts]
    return targets[continuous], origins[continuous]


def forecast_evaluation_contract(
    prediction_task: str | None, horizon_hours: object, *, legacy_task: str
) -> dict[str, object]:
    """Describe implemented timing; legacy row models never claim a 24h horizon."""

    if prediction_task not in (None, "", HISTORICAL_FORECAST_TASK):
        raise ValueError(f"Unsupported prediction_task: {prediction_task}")
    historical = prediction_task == HISTORICAL_FORECAST_TASK
    return {
        "task": HISTORICAL_FORECAST_TASK if historical else legacy_task,
        "information_set": (
            "observed_measurements_through_forecast_origin"
            if historical else "legacy_model_specific_observation_rows"
        ),
        "horizon_hours": validate_forecast_horizon(horizon_hours) if historical else None,
        "prediction_key": (
            ["timestamp", "plant_id", "forecast_origin", "horizon_hours"]
            if historical else ["timestamp", "plant_id"]
        ),
        "prediction_schema": "solar-forecast-prediction.v2" if historical else "solar-forecast-prediction.v1",
        "origin_definition": "target_timestamp_minus_horizon" if historical else None,
        "observation_availability_assumption": (
            "hourly_measurement_available_at_timestamp; publication_latency_not_simulated"
            if historical else None
        ),
    }
