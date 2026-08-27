from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd


BASE_WEATHER_FEATURES = [
    "temperature_c",
    "precipitation_mm",
    "sunshine_hours",
    "solar_irradiance_mj_m2",
    "wind_speed_mps",
    "humidity_pct",
    "total_cloud_cover_tenths",
    "low_mid_cloud_cover_tenths",
]

TIME_FEATURES = [
    "hour",
    "dayofweek",
    "month",
    "hour_sin",
    "hour_cos",
    "dayofyear_sin",
    "dayofyear_cos",
]

HISTORY_FEATURES = [
    "generation_lag_24h_mwh",
    "generation_lag_168h_mwh",
    "generation_rolling_7d_mean_mwh",
]

HISTORY_OBSERVATION_FEATURES = [
    "generation_lag_24h_observed",
    "generation_lag_168h_observed",
    "generation_rolling_7d_observation_ratio",
]

STATIC_FEATURES = [
    "capacity_mw",
    "tilt_deg",
    "station_latitude",
    "station_longitude",
    "station_elevation_m",
]

SOLAR_GEOMETRY_FEATURES = [
    "solar_elevation_sin",
    "clear_sky_irradiance_proxy",
    "is_daylight",
]

OBSERVATION_MASK_FEATURES = [
    "precipitation_observed",
    "sunshine_observed",
    "solar_irradiance_observed",
    "weather_missing_count",
]

CANDIDATE_MODEL_FEATURES = [
    *SOLAR_GEOMETRY_FEATURES,
    *OBSERVATION_MASK_FEATURES,
]

MISSINGNESS_CANDIDATE_FEATURES = [*HISTORY_OBSERVATION_FEATURES]

SELECTED_V2_MODEL_FEATURES = [
    *BASE_WEATHER_FEATURES,
    *TIME_FEATURES,
    *HISTORY_FEATURES,
    *STATIC_FEATURES,
]

SELECTED_MODEL_FEATURES = [*SELECTED_V2_MODEL_FEATURES, *SOLAR_GEOMETRY_FEATURES]

MODEL_READY_FEATURES = [
    *SELECTED_V2_MODEL_FEATURES,
    *CANDIDATE_MODEL_FEATURES,
    *MISSINGNESS_CANDIDATE_FEATURES,
]

EXCLUDED_DEFAULT_FEATURES = {
    "vapor_pressure_hpa": "humidity duplicates most of its information",
    "snow_depth_cm": "only 1.12% observed in the evaluated slice",
    "lowest_cloud_base_100m": "only 48.75% observed and did not improve MAE",
    "external_generation_forecast_mwh": "separate external baseline; never a direct default feature",
}


@dataclass(frozen=True)
class HistoryPolicy:
    forecast_horizon_hours: int = 24
    weekly_lag_hours: int = 168
    rolling_window_hours: int = 168
    rolling_min_observations: int = 24


class LeakageSafeFeatureEngineer:
    """Build time and history features available at a day-ahead issue time."""

    def __init__(self, policy: HistoryPolicy | None = None):
        self.policy = policy or HistoryPolicy()

    def transform(
        self,
        frame: pd.DataFrame,
        *,
        timestamp_column: str = "timestamp",
        entity_column: str = "plant_id",
        target_column: str = "generation_mwh",
    ) -> pd.DataFrame:
        required = {timestamp_column, entity_column, target_column}
        missing = required - set(frame.columns)
        if missing:
            raise ValueError(f"Feature engineering columns are missing: {sorted(missing)}")
        result = frame.copy()
        result[timestamp_column] = pd.to_datetime(result[timestamp_column], errors="coerce")
        if result[timestamp_column].isna().any():
            raise ValueError("Feature engineering input contains invalid timestamps")
        if result.duplicated([timestamp_column, entity_column]).any():
            raise ValueError("Feature engineering requires unique timestamp/entity keys")
        result = result.sort_values([entity_column, timestamp_column], kind="stable").reset_index(drop=True)

        original_weather = [column for column in BASE_WEATHER_FEATURES if column in result]
        result["weather_missing_count"] = (
            result[original_weather].isna().sum(axis=1) if original_weather else 0
        )
        for column, output in (
            ("precipitation_mm", "precipitation_observed"),
            ("sunshine_hours", "sunshine_observed"),
            ("solar_irradiance_mj_m2", "solar_irradiance_observed"),
        ):
            result[output] = result[column].notna().astype(int) if column in result else 0

        solar_elevation_sin, clear_sky_proxy = self._solar_geometry(result, timestamp_column)
        result["solar_elevation_sin"] = solar_elevation_sin
        result["clear_sky_irradiance_proxy"] = clear_sky_proxy
        result["is_daylight"] = solar_elevation_sin.gt(0).astype(int)

        # In the retained ASOS hourly exports, precipitation is event-like and
        # most dry hours are blank. Treat a blank as dry only while independent
        # core sensors are present, and retain the observation mask either way.
        if "precipitation_mm" in result:
            core = [
                column for column in ("temperature_c", "wind_speed_mps", "humidity_pct")
                if column in result
            ]
            core_available = result[core].notna().sum(axis=1).ge(2) if core else False
            dry_blank = result["precipitation_mm"].isna() & core_available
            result.loc[dry_blank, "precipitation_mm"] = 0.0

        night = result["is_daylight"].eq(0) & result["solar_elevation_sin"].notna()
        for column in ("sunshine_hours", "solar_irradiance_mj_m2"):
            if column in result:
                result.loc[night & result[column].isna(), column] = 0.0

        timestamp = result[timestamp_column]
        result["hour"] = timestamp.dt.hour
        result["dayofweek"] = timestamp.dt.dayofweek
        result["month"] = timestamp.dt.month
        result["hour_sin"] = np.sin(2 * np.pi * timestamp.dt.hour / 24)
        result["hour_cos"] = np.cos(2 * np.pi * timestamp.dt.hour / 24)
        result["dayofyear_sin"] = np.sin(2 * np.pi * timestamp.dt.dayofyear / 365.25)
        result["dayofyear_cos"] = np.cos(2 * np.pi * timestamp.dt.dayofyear / 365.25)

        for lag in (self.policy.forecast_horizon_hours, self.policy.weekly_lag_hours):
            lookup = result[[entity_column, timestamp_column, target_column]].copy()
            lookup[timestamp_column] += pd.to_timedelta(lag, unit="h")
            feature = f"generation_lag_{lag}h_mwh"
            lookup = lookup.rename(columns={target_column: feature})
            result = result.merge(
                lookup,
                on=[entity_column, timestamp_column],
                how="left",
                validate="one_to_one",
                sort=False,
            )
            result[f"generation_lag_{lag}h_observed"] = result[feature].notna().astype(int)

        rolling_parts: list[pd.DataFrame] = []
        for entity, group in result.groupby(entity_column, sort=False):
            series = group.set_index(timestamp_column)[target_column]
            hourly = series.reindex(pd.date_range(series.index.min(), series.index.max(), freq="h"))
            available = hourly.shift(self.policy.forecast_horizon_hours).notna().astype(float)
            rolling_source = hourly.shift(self.policy.forecast_horizon_hours)
            rolling = rolling_source.rolling(
                    self.policy.rolling_window_hours,
                    min_periods=self.policy.rolling_min_observations,
                ).mean()
            observation_ratio = available.rolling(
                self.policy.rolling_window_hours,
                min_periods=1,
            ).mean()
            rolling_parts.append(
                pd.DataFrame(
                    {
                        entity_column: entity,
                        timestamp_column: rolling.index,
                        "generation_rolling_7d_mean_mwh": rolling.to_numpy(),
                        "generation_rolling_7d_observation_ratio": observation_ratio.to_numpy(),
                    }
                )
            )
        rolling_frame = pd.concat(rolling_parts, ignore_index=True)
        result = result.merge(
            rolling_frame,
            on=[entity_column, timestamp_column],
            how="left",
            validate="one_to_one",
            sort=False,
        )
        return result.sort_values([timestamp_column, entity_column], kind="stable").reset_index(drop=True)

    @staticmethod
    def _solar_geometry(
        frame: pd.DataFrame, timestamp_column: str
    ) -> tuple[pd.Series, pd.Series]:
        """Approximate solar elevation and clear-sky horizontal potential for KST."""

        if not {"station_latitude", "station_longitude"}.issubset(frame.columns):
            missing = pd.Series(np.nan, index=frame.index, dtype=float)
            return missing, missing.copy()
        timestamp = frame[timestamp_column]
        latitude = np.deg2rad(pd.to_numeric(frame["station_latitude"], errors="coerce"))
        longitude = pd.to_numeric(frame["station_longitude"], errors="coerce")
        fractional_hour = timestamp.dt.hour + timestamp.dt.minute / 60
        gamma = 2 * np.pi / 365 * (timestamp.dt.dayofyear - 1 + (fractional_hour - 12) / 24)
        equation_of_time = 229.18 * (
            0.000075
            + 0.001868 * np.cos(gamma)
            - 0.032077 * np.sin(gamma)
            - 0.014615 * np.cos(2 * gamma)
            - 0.040849 * np.sin(2 * gamma)
        )
        declination = (
            0.006918
            - 0.399912 * np.cos(gamma)
            + 0.070257 * np.sin(gamma)
            - 0.006758 * np.cos(2 * gamma)
            + 0.000907 * np.sin(2 * gamma)
            - 0.002697 * np.cos(3 * gamma)
            + 0.00148 * np.sin(3 * gamma)
        )
        # KST is UTC+9 and has a standard meridian of 135 degrees east.
        true_solar_minutes = fractional_hour * 60 + equation_of_time + 4 * longitude - 60 * 9
        hour_angle = np.deg2rad(true_solar_minutes / 4 - 180)
        cos_zenith = (
            np.sin(latitude) * np.sin(declination)
            + np.cos(latitude) * np.cos(declination) * np.cos(hour_angle)
        ).clip(-1, 1)
        elevation_sin = pd.Series(cos_zenith, index=frame.index).where(latitude.notna() & longitude.notna())
        eccentricity = 1.00011 + 0.034221 * np.cos(gamma) + 0.00128 * np.sin(gamma)
        clear_sky_proxy = elevation_sin.clip(lower=0) * eccentricity
        return elevation_sin, clear_sky_proxy
