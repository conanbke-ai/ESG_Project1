from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd

from solar_forecast.collectors.normalization import read_csv_with_fallback


WEATHER_COLUMN_MAP = {
    "지점": "station_id",
    "지점명": "station_name",
    "일시": "timestamp",
    "기온(°C)": "temperature_c",
    "강수량(mm)": "precipitation_mm",
    "풍속(m/s)": "wind_speed_mps",
    "습도(%)": "humidity_pct",
    "일조(hr)": "sunshine_hours",
    "일사(MJ/m2)": "solar_irradiance_mj_m2",
    "전운량(10분위)": "total_cloud_cover_tenths",
    "중하층운량(10분위)": "low_mid_cloud_cover_tenths",
}

WEATHER_COLUMNS = list(WEATHER_COLUMN_MAP.values()) + [
    "station_latitude",
    "station_longitude",
    "station_elevation_m",
]


class KmaAsosNormalizer:
    """Normalize selected ASOS fields and join stable station coordinates."""

    def __init__(self, station_metadata_path: Path):
        self.station_metadata_path = Path(station_metadata_path)

    def read(self, paths: Iterable[Path]) -> pd.DataFrame:
        frames = [read_csv_with_fallback(Path(path)) for path in paths]
        if not frames:
            raise ValueError("At least one KMA hourly file is required")
        return self.transform(pd.concat(frames, ignore_index=True))

    def transform(self, frame: pd.DataFrame) -> pd.DataFrame:
        source = frame.copy()
        source.columns = [str(column).strip() for column in source.columns]
        missing = set(WEATHER_COLUMN_MAP) - set(source.columns)
        if missing:
            raise ValueError(f"KMA ASOS columns are missing: {sorted(missing)}")
        result = source[list(WEATHER_COLUMN_MAP)].rename(columns=WEATHER_COLUMN_MAP)
        result["timestamp"] = pd.to_datetime(result["timestamp"], errors="coerce")
        numeric = [column for column in result.columns if column not in {"station_name", "timestamp"}]
        for column in numeric:
            result[column] = pd.to_numeric(result[column], errors="coerce")
        result = result.dropna(subset=["station_id", "timestamp"])
        result["station_id"] = result["station_id"].astype(int)
        result = result.drop_duplicates(["station_id", "timestamp"], keep="last")

        metadata = read_csv_with_fallback(self.station_metadata_path)
        metadata.columns = [str(column).strip() for column in metadata.columns]
        required_metadata = {"지점", "위도", "경도", "노장해발고도(m)"}
        missing_metadata = required_metadata - set(metadata.columns)
        if missing_metadata:
            raise ValueError(f"KMA station metadata columns are missing: {sorted(missing_metadata)}")
        if "시작일" in metadata.columns:
            metadata["시작일"] = pd.to_datetime(metadata["시작일"], errors="coerce")
            metadata = metadata.sort_values("시작일", kind="stable")
        metadata = metadata.drop_duplicates("지점", keep="last")
        metadata = metadata[["지점", "위도", "경도", "노장해발고도(m)"]].rename(
            columns={
                "지점": "station_id",
                "위도": "station_latitude",
                "경도": "station_longitude",
                "노장해발고도(m)": "station_elevation_m",
            }
        )
        for column in metadata.columns:
            metadata[column] = pd.to_numeric(metadata[column], errors="coerce")
        metadata = metadata.dropna(subset=["station_id"])
        metadata["station_id"] = metadata["station_id"].astype(int)
        result = result.merge(metadata, on="station_id", how="left", validate="many_to_one")
        return result[WEATHER_COLUMNS].sort_values(["timestamp", "station_id"]).reset_index(drop=True)
