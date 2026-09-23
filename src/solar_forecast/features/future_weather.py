"""Leakage-safe, unit-aligned future-weather covariates for solar forecasts."""
from __future__ import annotations

from pathlib import Path
from typing import Sequence

import pandas as pd


FUTURE_WEATHER_CONTRACT = "solar-future-weather-forecast.v2"
SUPPORTED_FIXED_LEADS = {24, 72}

# These variables deliberately mirror the physical meaning and units of the
# historical ASOS inputs as closely as the forecast provider allows.
FUTURE_WEATHER_CORE_FEATURES = (
    "future_temperature_c",
    "future_humidity_pct",
    "future_precipitation_mm",
    "future_total_cloud_cover_tenths",
    "future_wind_speed_mps",
    "future_solar_irradiance_mj_m2",
    "future_sunshine_hours",
)

# DNI/DHI do not have direct ASOS counterparts in the current Gold contract.
# Keep them available for an explicit ablation instead of silently changing the
# default information set.
FUTURE_WEATHER_RADIATION_COMPONENT_FEATURES = (
    "future_dni_mj_m2",
    "future_dhi_mj_m2",
)

FUTURE_WEATHER_FEATURE_PROFILES = {
    "aligned_core": FUTURE_WEATHER_CORE_FEATURES,
    "aligned_core_plus_components": (
        *FUTURE_WEATHER_CORE_FEATURES,
        *FUTURE_WEATHER_RADIATION_COMPONENT_FEATURES,
    ),
}

# Backward-compatible public name: the default model path uses aligned_core.
FUTURE_WEATHER_FEATURES = FUTURE_WEATHER_CORE_FEATURES
FUTURE_WEATHER_ARCHIVE_FEATURES = FUTURE_WEATHER_FEATURE_PROFILES[
    "aligned_core_plus_components"
]

FUTURE_WEATHER_KEYS = (
    "plant_id",
    "timestamp",
    "forecast_origin",
    "horizon_hours",
)
FUTURE_WEATHER_COORDINATE_COLUMNS = (
    "weather_query_latitude",
    "weather_query_longitude",
    "grid_latitude",
    "grid_longitude",
    "coordinate_source",
    "cell_selection",
)
ALLOWED_COORDINATE_SOURCES = {
    "plant_registry_coordinates",
    "reviewed_asos_station_proxy",
}


def future_weather_feature_columns(profile: str = "aligned_core") -> tuple[str, ...]:
    """Return one explicit future-weather feature contract."""

    try:
        return tuple(FUTURE_WEATHER_FEATURE_PROFILES[str(profile)])
    except KeyError as exc:
        raise ValueError(
            "Unknown future-weather feature profile "
            f"{profile!r}; expected one of {sorted(FUTURE_WEATHER_FEATURE_PROFILES)}"
        ) from exc


def read_future_weather(
    source: str | Path,
    *,
    horizon_hours: int,
    plant_ids: Sequence[str] | None = None,
    feature_profile: str = "aligned_core",
) -> pd.DataFrame:
    """Read one fixed-lead archived forecast cohort and fail closed on leakage."""

    horizon = int(horizon_hours)
    if horizon not in SUPPORTED_FIXED_LEADS:
        raise ValueError(
            "Archived future-weather covariates currently support only fixed "
            f"24h/72h leads, got {horizon}h"
        )
    selected_features = future_weather_feature_columns(feature_profile)
    path = Path(source)
    if not path.is_absolute():
        from solar_forecast.config_loader import PROJECT_ROOT

        path = PROJECT_ROOT / path
    if not path.is_file():
        raise FileNotFoundError(
            f"Future-weather archive is missing: {path}. "
            "Collect archived forecasts before enabling future covariates."
        )

    required = [
        *FUTURE_WEATHER_KEYS,
        *selected_features,
        *FUTURE_WEATHER_COORDINATE_COLUMNS,
        "forecast_model",
        "forecast_source",
    ]
    frame = pd.read_csv(
        path,
        usecols=lambda column: column in required,
        dtype={"plant_id": str},
        low_memory=False,
    )
    missing = set(required) - set(frame)
    if missing:
        raise ValueError(
            "Future-weather archive is not compatible with "
            f"{FUTURE_WEATHER_CONTRACT}: missing {sorted(missing)}. "
            "Rebuild it with tools/collect_future_weather_previous_runs.py."
        )

    frame["timestamp"] = pd.to_datetime(frame["timestamp"], errors="raise")
    frame["forecast_origin"] = pd.to_datetime(frame["forecast_origin"], errors="raise")
    frame["horizon_hours"] = pd.to_numeric(
        frame["horizon_hours"], errors="raise"
    ).astype(int)
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

    for column, lower, upper in (
        ("weather_query_latitude", 32.0, 39.5),
        ("weather_query_longitude", 124.0, 132.5),
        ("grid_latitude", -90.0, 90.0),
        ("grid_longitude", -180.0, 180.0),
    ):
        frame[column] = pd.to_numeric(frame[column], errors="coerce")
        if frame[column].isna().any() or not frame[column].between(lower, upper).all():
            raise ValueError(f"Future-weather archive has invalid {column}")
    invalid_sources = set(
        frame["coordinate_source"].dropna().astype(str).unique()
    ) - ALLOWED_COORDINATE_SOURCES
    if invalid_sources:
        raise ValueError(
            "Future-weather archive uses an unapproved spatial proxy: "
            f"{sorted(invalid_sources)}"
        )
    if not frame["cell_selection"].eq("nearest").all():
        raise ValueError("Future-weather archive must use nearest grid selection")

    for column in selected_features:
        frame[column] = pd.to_numeric(frame[column], errors="coerce")

    all_missing_rows = frame[list(selected_features)].isna().all(axis=1)
    entirely_missing_features = [
        column
        for column in selected_features
        if frame[column].isna().all()
    ]

    frame.attrs["future_weather_contract"] = FUTURE_WEATHER_CONTRACT
    frame.attrs["all_missing_rows"] = int(all_missing_rows.sum())
    frame.attrs["all_missing_fraction"] = (
        float(all_missing_rows.mean()) if len(frame) else 0.0
    )
    frame.attrs["entirely_missing_features"] = entirely_missing_features
    frame.attrs["feature_profile"] = feature_profile
    frame.attrs["feature_columns"] = list(selected_features)
    return frame.sort_values(
        ["timestamp", "plant_id"], kind="stable"
    ).reset_index(drop=True)


def merge_future_weather(
    forecast_samples: pd.DataFrame,
    future_weather: pd.DataFrame,
    *,
    minimum_coverage: float = 0.98,
    feature_columns: Sequence[str] | None = None,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """Attach fixed-lead weather to target rows without changing forecast keys."""

    missing = set(FUTURE_WEATHER_KEYS) - set(forecast_samples)
    if missing:
        raise ValueError(f"Forecast samples omit future-weather keys: {sorted(missing)}")
    if not 0 < minimum_coverage <= 1:
        raise ValueError("minimum_coverage must be in (0, 1]")

    selected = tuple(
        feature_columns
        or future_weather.attrs.get("feature_columns")
        or FUTURE_WEATHER_CORE_FEATURES
    )
    missing_features = set(selected) - set(future_weather.columns)
    if missing_features:
        raise ValueError(
            f"Future-weather frame omits selected features: {sorted(missing_features)}"
        )

    sidecar = future_weather[
        [
            *FUTURE_WEATHER_KEYS,
            *selected,
            *FUTURE_WEATHER_COORDINATE_COLUMNS,
            "forecast_model",
            "forecast_source",
        ]
    ].copy()
    merged = forecast_samples.merge(
        sidecar,
        on=list(FUTURE_WEATHER_KEYS),
        how="left",
        validate="one_to_one",
        sort=False,
    )
    available = ~merged[list(selected)].isna().all(axis=1)
    coverage = float(available.mean()) if len(merged) else 0.0
    if coverage < minimum_coverage:
        raise ValueError(
            f"Future-weather coverage {coverage:.2%} is below "
            f"{minimum_coverage:.2%}"
        )
    merged = merged.loc[available].copy()
    evidence = {
        "contract": FUTURE_WEATHER_CONTRACT,
        "feature_profile": future_weather.attrs.get(
            "feature_profile", "aligned_core"
        ),
        "rows_before": int(len(forecast_samples)),
        "rows_after": int(len(merged)),
        "coverage": coverage,
        "feature_columns": list(selected),
        "forecast_models": sorted(
            merged["forecast_model"].dropna().astype(str).unique().tolist()
        ),
        "forecast_sources": sorted(
            merged["forecast_source"].dropna().astype(str).unique().tolist()
        ),
        "leakage_guard": "forecast_origin_equals_target_minus_fixed_lead",
        "unit_alignment": {
            "cloud_cover": "percent_divided_by_10_to_ASOS_tenths",
            "shortwave_radiation": "hourly_mean_W_m2_times_0.0036_to_MJ_m2",
            "sunshine_duration": "seconds_divided_by_3600_to_hours",
        },
        "spatial_contract": (
            "plant_coordinates_else_reviewed_asos_proxy; "
            "region_centroids_forbidden"
        ),
        "coordinate_sources": sorted(
            merged["coordinate_source"].dropna().astype(str).unique().tolist()
        ),
        "cell_selection": sorted(
            merged["cell_selection"].dropna().astype(str).unique().tolist()
        ),
    }
    return merged, evidence


def attach_future_weather_to_observations(
    observations: pd.DataFrame,
    future_weather: pd.DataFrame,
    *,
    entity_column: str = "plant_id",
    timestamp_column: str = "timestamp",
    minimum_coverage: float = 0.98,
    feature_columns: Sequence[str] | None = None,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """Attach target-time forecast values for the CNN future-covariate branch."""

    selected = tuple(
        feature_columns
        or future_weather.attrs.get("feature_columns")
        or FUTURE_WEATHER_CORE_FEATURES
    )
    if entity_column != "plant_id":
        sidecar = future_weather.rename(columns={"plant_id": entity_column})
    else:
        sidecar = future_weather
    keys = [entity_column, timestamp_column]
    side = sidecar.rename(columns={"timestamp": timestamp_column})
    if side.duplicated(keys).any():
        raise ValueError("Future-weather archive is not unique per plant target time")
    merged = observations.merge(
        side[
            [
                *keys,
                *selected,
                *FUTURE_WEATHER_COORDINATE_COLUMNS,
                "forecast_origin",
                "horizon_hours",
            ]
        ],
        on=keys,
        how="left",
        validate="one_to_one",
        sort=False,
    )
    available = ~merged[list(selected)].isna().all(axis=1)
    evidence = {
        "contract": FUTURE_WEATHER_CONTRACT,
        "feature_profile": future_weather.attrs.get(
            "feature_profile", "aligned_core"
        ),
        "row_weather_coverage": float(available.mean()) if len(merged) else 0.0,
        "minimum_forecast_cohort_coverage": float(minimum_coverage),
        "feature_columns": list(selected),
        "spatial_contract": (
            "plant_coordinates_else_reviewed_asos_proxy; "
            "region_centroids_forbidden"
        ),
        "coordinate_sources": sorted(
            merged["coordinate_source"].dropna().astype(str).unique().tolist()
        ),
    }
    return merged, evidence
