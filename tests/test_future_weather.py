"""Future-weather contracts keep fixed lead time and plant-coordinate provenance."""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from solar_forecast.features.future_weather import (
    FUTURE_WEATHER_FEATURES,
    merge_future_weather,
    read_future_weather,
)


def _archive(path: Path, *, coordinate_source: str = "plant_registry_coordinates") -> Path:
    frame = pd.DataFrame(
        {
            "plant_id": ["ewp:울산화력태양광"],
            "timestamp": ["2024-07-02T12:00:00"],
            "forecast_origin": ["2024-07-01T12:00:00"],
            "horizon_hours": [24],
            "plant_latitude": [35.47765],
            "plant_longitude": [129.3808],
            "grid_latitude": [35.5],
            "grid_longitude": [129.4],
            "coordinate_source": [coordinate_source],
            "cell_selection": ["nearest"],
            "forecast_model": ["jma_msm"],
            "forecast_source": ["open_meteo_previous_runs"],
            **{
                name: [value]
                for name, value in zip(
                    FUTURE_WEATHER_FEATURES,
                    (26.0, 62.0, 0.0, 35.0, 2.1, 620.0, 510.0, 110.0),
                    strict=True,
                )
            },
        }
    )
    frame.to_csv(path, index=False)
    return path


def test_future_weather_accepts_only_fixed_lead_plant_coordinates(tmp_path: Path):
    path = _archive(tmp_path / "weather.csv")
    frame = read_future_weather(path, horizon_hours=24)
    assert frame.iloc[0]["coordinate_source"] == "plant_registry_coordinates"
    assert frame.iloc[0]["plant_latitude"] == pytest.approx(35.47765)
    assert frame.iloc[0]["plant_longitude"] == pytest.approx(129.3808)


def test_future_weather_rejects_region_or_station_coordinate_fallback(tmp_path: Path):
    path = _archive(
        tmp_path / "weather.csv",
        coordinate_source="asos_station_coordinates",
    )
    with pytest.raises(ValueError, match="plant_registry_coordinates"):
        read_future_weather(path, horizon_hours=24)


def test_future_weather_merge_preserves_spatial_provenance(tmp_path: Path):
    archive = read_future_weather(_archive(tmp_path / "weather.csv"), horizon_hours=24)
    samples = pd.DataFrame(
        {
            "plant_id": ["ewp:울산화력태양광"],
            "timestamp": [pd.Timestamp("2024-07-02 12:00:00")],
            "forecast_origin": [pd.Timestamp("2024-07-01 12:00:00")],
            "horizon_hours": [24],
            "y_true": [0.31],
        }
    )
    merged, evidence = merge_future_weather(
        samples,
        archive,
        minimum_coverage=1.0,
    )
    assert len(merged) == 1
    assert evidence["spatial_contract"] == "plant_registry_coordinates_only"
    assert evidence["coordinate_sources"] == ["plant_registry_coordinates"]
