"""Leakage-safe future-weather covariates for historical solar forecasts."""
from __future__ import annotations

from pathlib import Path
from typing import Sequence

import numpy as np
import pandas as pd


FUTURE_WEATHER_CONTRACT = "solar-future-weather-forecast.v1"
SUPPORTED_FIXED_LEADS = {24, 72}
FUTURE_WEATHER_FEATURES = (
    "future_temperature_c",
    "future_humidity_pct",
    "future_precipitation_mm",
    "future_cloud_cover_pct",
    "future_wind_speed_mps",
    "future_ghi_w_m2",
    "future_dni_w_m2",
    "future_dhi_w_m2",
)
FUTURE_WEATHER_KEYS = (
    "plant_id",
    "timestamp",
    "forecast_origin",
    "horizon_hours",
)


def read_future_weather(
    source: str | Path,
    *,
    horizon_hours: int,
    plant_ids: Sequence[str] | None = None,
) -> pd.DataFrame:
    """Read one fixed-lead archived forecast cohort and fail closed on leakage.

    The sidecar must contain weather values that were forecasts at the declared
    forecast origin. It is never acceptable to substitute future observations
    into these columns for benchmark scoring.
    """

    horizon = int(horizon_hours)
    if horizon not in SUPPORTED_FIXED_LEADS:
        raise ValueError(
            "Archived future-weather covariates currently support only fixed "
            f"24h/72h leads, got {horizon}h"
        )
    path = Path(source)
    if not path.is_absolute():
        from solar_forecast.config_loader import PROJECT_ROOT

        path = PROJECT_ROOT / path
    if not path.is_file():
        raise FileNotFoundError(
            f"Future-weather archive is missing: {path}. "
            "Collect archived forecasts before enabling future covariates."
        )

    required = [*FUTURE_WEATHER_KEYS, *FUTURE_WEATHER_FEATURES, "forecast_model", "forecast_source"]
    frame = pd.read_csv(
        path,
        usecols=lambda column: column in required,
        dtype={"plant_id": str},
        low_memory=False,
    )
    missing = set(required) - set(frame)
    if missing:
        raise ValueError(f"Future-weather archive columns are missing: {sorted(missing)}")

    frame["timestamp"] = pd.to_datetime(frame["timestamp"], errors="raise")
    frame["forecast_origin"] = pd.to_datetime(frame["forecast_origin"], errors="raise")
    frame["horizon_hours"] = pd.to_numeric(frame["horizon_hours"], errors="raise").astype(int)
    frame = frame.loc[frame["horizon_hours"].eq(horizon)].copy()
    if plant_ids is not None:
        allowed = {str(value) for value in plant_ids}
        frame = frame.loc[frame["plant_id"].astype(str).isin(allowed)].copy()
    if frame.empty:
        raise ValueError(f"Future-weather archive has no rows for {horizon}h")

    expected_origin = frame["timestamp"] - pd.to_timedelta(horizon, unit="h")
    if not expected_origin.eq(frame["forecast_origin"]).all():
        raise ValueError(
            "Future-weather forecast_origin must equal target timestamp minus "
            "the declared fixed lead; future observations or mismatched runs "
            "cannot enter the benchmark."
        )
    if frame.duplicated(list(FUTURE_WEATHER_KEYS)).any():
        raise ValueError("Future-weather archive contains duplicate forecast keys")

    for column in FUTURE_WEATHER_FEATURES:
        frame[column] = pd.to_numeric(frame[column], errors="coerce")
    # Missing forecasts are preserved for model-specific missing-value handling,
    # but an entirely absent weather vector is not a valid forecast record.
    if frame[list(FUTURE_WEATHER_FEATURES)].isna().all(axis=1).any():
        raise ValueError("Future-weather archive contains rows with no forecast variables")

    return frame.sort_values(["timestamp", "plant_id"], kind="stable").reset_index(drop=True)


def merge_future_weather(
    forecast_samples: pd.DataFrame,
    future_weather: pd.DataFrame,
    *,
    minimum_coverage: float = 0.98,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """Attach fixed-lead weather to target rows without changing forecast keys."""

    missing = set(FUTURE_WEATHER_KEYS) - set(forecast_samples)
    if missing:
        raise ValueError(f"Forecast samples omit future-weather keys: {sorted(missing)}")
    if not 0 < minimum_coverage <= 1:
        raise ValueError("minimum_coverage must be in (0, 1]")

    sidecar = future_weather[
        [*FUTURE_WEATHER_KEYS, *FUTURE_WEATHER_FEATURES, "forecast_model", "forecast_source"]
    ].copy()
    merged = forecast_samples.merge(
        sidecar,
        on=list(FUTURE_WEATHER_KEYS),
        how="left",
        validate="one_to_one",
        sort=False,
    )
    available = ~merged[list(FUTURE_WEATHER_FEATURES)].isna().all(axis=1)
    coverage = float(available.mean()) if len(merged) else 0.0
    if coverage < minimum_coverage:
        raise ValueError(
            f"Future-weather coverage {coverage:.2%} is below "
            f"{minimum_coverage:.2%}"
        )
    # Keep partial-variable gaps as NaN so imputers fit on Train only. Rows with
    # no archived forecast are removed identically for every candidate using
    # this future-weather contract.
    merged = merged.loc[available].copy()
    evidence = {
        "contract": FUTURE_WEATHER_CONTRACT,
        "rows_before": int(len(forecast_samples)),
        "rows_after": int(len(merged)),
        "coverage": coverage,
        "feature_columns": list(FUTURE_WEATHER_FEATURES),
        "forecast_models": sorted(
            merged["forecast_model"].dropna().astype(str).unique().tolist()
        ),
        "forecast_sources": sorted(
            merged["forecast_source"].dropna().astype(str).unique().tolist()
        ),
        "leakage_guard": "forecast_origin_equals_target_minus_fixed_lead",
    }
    return merged, evidence


def attach_future_weather_to_observations(
    observations: pd.DataFrame,
    future_weather: pd.DataFrame,
    *,
    entity_column: str = "plant_id",
    timestamp_column: str = "timestamp",
    minimum_coverage: float = 0.98,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """Attach target-time forecast values to raw rows for sequence models.

    These columns are consumed only at a target position by the CNN future
    covariate branch; they must never be inserted into the historical input
    window.
    """

    if entity_column != "plant_id":
        sidecar = future_weather.rename(columns={"plant_id": entity_column})
    else:
        sidecar = future_weather
    keys = [entity_column, timestamp_column]
    side = sidecar.rename(columns={"timestamp": timestamp_column})
    if side.duplicated(keys).any():
        raise ValueError("Future-weather archive is not unique per plant target time")
    merged = observations.merge(
        side[[*keys, *FUTURE_WEATHER_FEATURES, "forecast_origin", "horizon_hours"]],
        on=keys,
        how="left",
        validate="one_to_one",
        sort=False,
    )
    available = ~merged[list(FUTURE_WEATHER_FEATURES)].isna().all(axis=1)
    # Coverage is evaluated only across rows that can plausibly become target
    # samples later. Early history may predate the forecast archive, so caller
    # additionally checks the actual forecast cohort after window construction.
    evidence = {
        "contract": FUTURE_WEATHER_CONTRACT,
        "row_weather_coverage": float(available.mean()) if len(merged) else 0.0,
        "minimum_forecast_cohort_coverage": float(minimum_coverage),
        "feature_columns": list(FUTURE_WEATHER_FEATURES),
    }
    return merged, evidence
