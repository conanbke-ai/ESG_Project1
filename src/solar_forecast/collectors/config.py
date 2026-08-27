from __future__ import annotations

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
    output_dir: Path = Path("file/raw")
    existing_weather_dir: Path = Path("file/KMA_data_file")
    overwrite: bool = False

    def __post_init__(self) -> None:
        if self.end_date < self.start_date:
            raise ValueError("end_date must be on or after start_date")


def load_source_catalog() -> dict[str, dict[str, Any]]:
    resource = files("solar_forecast.collectors").joinpath("sources.json")
    return json.loads(resource.read_text(encoding="utf-8"))
