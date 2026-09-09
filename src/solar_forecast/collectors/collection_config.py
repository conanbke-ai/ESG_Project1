"""수집 기간·공급자·저장 경로·ASOS 모드·호출 한도 검증."""
from __future__ import annotations

from solar_forecast.infrastructure.project_paths import BRONZE_ROOT, COLLECTOR_SILVER_ROOT, WEATHER_ARCHIVE_ROOT

from dataclasses import dataclass, field
from datetime import date
from importlib.resources import files
import json
from pathlib import Path
from typing import Any, Sequence


@dataclass(frozen=True)
class CollectionConfig:
    start_date: date
    end_date: date
    station_ids: Sequence[str] = field(default_factory=tuple)
    sources: Sequence[str] = ("koen", "kospo", "ewp", "iwest", "kma")
    output_dir: Path = BRONZE_ROOT
    standardized_output_dir: Path = COLLECTOR_SILVER_ROOT
    existing_weather_dir: Path = WEATHER_ARCHIVE_ROOT
    overwrite: bool = False
    komipo_station_codes: Sequence[str] = field(default_factory=tuple)
    api_max_calls: int = 900
    download_date: date = field(default_factory=date.today)
    kma_mode: str = "auto"

    def __post_init__(self) -> None:
        if self.end_date < self.start_date:
            raise ValueError("end_date must be on or after start_date")
        if self.api_max_calls < 1:
            raise ValueError("api_max_calls must be positive")
        if self.kma_mode not in {"auto", "api", "browser"}:
            raise ValueError("kma_mode must be auto, api, or browser")
        if any(not str(station).isdigit() or not 1 <= int(station) <= 999 for station in self.station_ids):
            raise ValueError("station_ids must contain numeric ASOS station IDs from 1 to 999")


def load_source_catalog() -> dict[str, dict[str, Any]]:
    resource = files("solar_forecast.collectors").joinpath("sources.json")
    return json.loads(resource.read_text(encoding="utf-8"))
