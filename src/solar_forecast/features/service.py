from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
from typing import Iterable

import pandas as pd

from solar_forecast.collectors.metadata import PlantMetadataCatalog
from solar_forecast.collectors.normalization import (
    GENERATION_COLUMNS,
    classify_energy_source,
    read_csv_with_fallback,
)
from solar_forecast.quality.policy import GenerationQualityPolicy, QUALITY_COLUMNS

from .engineering import (
    MODEL_READY_FEATURES,
    SELECTED_MODEL_FEATURES,
    LeakageSafeFeatureEngineer,
)
from .weather import KmaAsosNormalizer


@dataclass(frozen=True)
class ModelDatasetResult:
    path: Path
    manifest_path: Path
    rows: int
    start: str
    end: str
    plants: int


class LegacyModelDatasetBuilder:
    """Upgrade the retained Korean merged file to the selected common feature contract."""

    source_columns = {
        "일시": "timestamp",
        "발전구분": "plant",
        "지역": "region",
        "지점번호": "station_id",
        "합산발전량(MWh)": "generation_mwh",
    }

    def __init__(
        self,
        weather_root: Path,
        metadata: PlantMetadataCatalog,
        engineer: LeakageSafeFeatureEngineer | None = None,
        quality_policy: GenerationQualityPolicy | None = None,
    ):
        self.weather_root = Path(weather_root)
        self.metadata = metadata
        self.engineer = engineer or LeakageSafeFeatureEngineer()
        self.quality_policy = quality_policy or GenerationQualityPolicy()

    def read_legacy_generation(self, source_path: Path) -> pd.DataFrame:
        """Read the retained merge only as a plant/station mapping and audit source."""

        raw = read_csv_with_fallback(Path(source_path))
        missing = set(self.source_columns) - set(raw.columns)
        if missing:
            raise ValueError(f"Merged source columns are missing: {sorted(missing)}")
        generation = raw[list(self.source_columns)].rename(columns=self.source_columns)
        generation["timestamp"] = pd.to_datetime(generation["timestamp"], errors="coerce")
        generation["generation_mwh"] = pd.to_numeric(generation["generation_mwh"], errors="coerce")
        generation["station_id"] = pd.to_numeric(generation["station_id"], errors="coerce")
        generation = generation.dropna(subset=["timestamp", "generation_mwh", "station_id"])
        generation["station_id"] = generation["station_id"].astype(int)
        generation["company"] = "kospo"
        generation["unit"] = ""
        classified = generation["plant"].map(classify_energy_source)
        generation["energy_source"] = classified.where(classified.ne("unknown"), "solar")
        generation["plant_id"] = "kospo:" + generation["plant"].astype(str).str.strip()
        for column in ("capacity_mw", "tilt_deg", "latitude", "longitude", "address"):
            generation[column] = pd.NA
        generation["source_file"] = Path(source_path).name
        return self.metadata.enrich(
            generation[GENERATION_COLUMNS + ["region", "station_id"]], aggregate=True
        )

    def build(
        self,
        source_path: Path,
        destination: Path,
        *,
        generation_paths: Iterable[Path] | None = None,
    ) -> ModelDatasetResult:
        legacy_generation = self.read_legacy_generation(source_path)
        generation = (
            self._from_standardized_generation(legacy_generation, generation_paths)
            if generation_paths is not None
            else legacy_generation
        )

        years = sorted(generation["timestamp"].dt.year.unique())
        weather_paths = [self.weather_root / f"OBS_ASOS_TIM_{year}.csv" for year in years]
        missing_weather = [path for path in weather_paths if not path.exists()]
        if missing_weather:
            raise FileNotFoundError(f"KMA hourly files are missing: {missing_weather}")
        weather = KmaAsosNormalizer(self.weather_root / "META_관측지점정보.csv").read(weather_paths)
        result = generation.merge(
            weather.drop(columns="station_name"),
            on=["station_id", "timestamp"],
            how="left",
            validate="many_to_one",
        )
        result = self.engineer.transform(result)
        result = self.quality_policy.apply(result)
        # Keep genuine gaps and cold-start rows. Dropping incomplete history
        # would preferentially retain healthy sensors/plants and bias training.
        context = [
            "timestamp", "company", "plant_id", "plant", "region", "station_id", "energy_source"
        ]
        result = result[[*context, *MODEL_READY_FEATURES, *QUALITY_COLUMNS, "generation_mwh"]]
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        result.to_csv(destination, index=False, encoding="utf-8-sig")

        coverage = {
            column: float(result[column].notna().mean()) for column in SELECTED_MODEL_FEATURES
        }
        manifest = {
            "created_at": datetime.now().isoformat(),
            "source": str(source_path),
            "generation_source": (
                "official_standardized_partitions"
                if generation_paths is not None
                else "legacy_merged_target"
            ),
            "legacy_source_usage": "plant-to-region/station mapping only",
            "dataset": str(destination),
            "target": "generation_mwh",
            "features": SELECTED_MODEL_FEATURES,
            "promoted_features": [
                "solar_elevation_sin", "clear_sky_irradiance_proxy", "is_daylight"
            ],
            "feature_selection_evidence": {
                "protocol": "3-fold purged expanding rolling-origin",
                "gap_hours": 168,
                "validation_window_hours": 2160,
                "reserved_calibration_fraction": 0.10,
                "reserved_test_fraction": 0.15,
                "energy_source": "solar only; hydro excluded",
                "generation_source": "official standardized raw aggregation",
                "baseline_23_mean_mae": 0.04028393219503514,
                "selected_26_mean_mae": 0.03903863517315129,
                "relative_mean_mae_improvement": 0.03091299185174775,
                "rejected_default_candidates": [
                    "weather observation masks (30 features)",
                    "history availability features (33 features)"
                ],
                "test_usage": "none; Calibration and Test reserved",
            },
            "history_rule": "lags are >=24h; rolling 7d is shifted by 24h",
            "history_missing_policy": (
                "preserve NaN and availability features; never delete a row only because history is missing"
            ),
            "energy_source_filter": "not applied in the retained model-ready table; model jobs filter solar",
            "quality_columns": QUALITY_COLUMNS,
            "quality_policy": {
                "negative_generation": "invalid; preserve flag and exclude from training, never replace with zero",
                "contextual_anomalies": "flag only; do not automatically impute or delete",
                "invalid_weather": "set missing with reason flag; impute from training data only",
            },
            "rows": len(result),
            "plants": int(result["plant_id"].nunique()),
            "start": result["timestamp"].min().isoformat(),
            "end": result["timestamp"].max().isoformat(),
            "coverage": coverage,
        }
        manifest_path = destination.with_name("model_ready_manifest.json")
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        return ModelDatasetResult(
            destination,
            manifest_path,
            len(result),
            manifest["start"],
            manifest["end"],
            manifest["plants"],
        )

    def _from_standardized_generation(
        self,
        legacy: pd.DataFrame,
        paths: Iterable[Path],
    ) -> pd.DataFrame:
        plants = set(legacy["plant"].astype(str))
        start, end = legacy["timestamp"].min(), legacy["timestamp"].max()
        parts: list[pd.DataFrame] = []
        for path in paths:
            source = pd.read_csv(
                Path(path),
                usecols=["timestamp", "company", "plant", "energy_source", "generation_mwh"],
            )
            source["timestamp"] = pd.to_datetime(source["timestamp"], errors="coerce")
            source = source.loc[
                source["plant"].astype(str).isin(plants)
                & source["timestamp"].between(start, end)
            ]
            if not source.empty:
                parts.append(source)
        if not parts:
            raise ValueError("No standardized official generation rows match the legacy plant/time mapping")
        generation = pd.concat(parts, ignore_index=True)
        generation["generation_mwh"] = pd.to_numeric(generation["generation_mwh"], errors="coerce")
        generation = generation.dropna(subset=["timestamp", "generation_mwh"])
        generation = generation.groupby(
            ["timestamp", "company", "plant", "energy_source"],
            as_index=False,
            sort=False,
        )["generation_mwh"].sum()

        mapping = legacy[["plant", "region", "station_id"]].drop_duplicates()
        ambiguous = mapping.groupby("plant").size()
        ambiguous = ambiguous[ambiguous.gt(1)]
        if not ambiguous.empty:
            raise ValueError(f"Legacy source has ambiguous plant/station mappings: {ambiguous.to_dict()}")
        generation = generation.merge(mapping, on="plant", how="inner", validate="many_to_one")
        generation["unit"] = ""
        generation["plant_id"] = generation["company"] + ":" + generation["plant"].astype(str)
        for column in ("capacity_mw", "tilt_deg", "latitude", "longitude", "address"):
            generation[column] = pd.NA
        generation["source_file"] = "official_standardized_partitions"
        generation = self.metadata.enrich(
            generation[GENERATION_COLUMNS + ["region", "station_id"]], aggregate=True
        )
        if generation.duplicated(["timestamp", "plant_id"]).any():
            raise ValueError("Official aggregate contains duplicate timestamp/plant_id keys")
        return generation.sort_values(["timestamp", "plant_id"], kind="stable").reset_index(drop=True)
