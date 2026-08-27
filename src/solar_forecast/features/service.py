from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
import re
import shutil
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
from .registry import KmaStationCatalog, NationwidePlantRegistryBuilder
from .weather import KmaAsosNormalizer


@dataclass(frozen=True)
class ModelDatasetResult:
    path: Path
    manifest_path: Path
    rows: int
    start: str
    end: str
    plants: int
    registry_path: Path | None = None
    quarantined_plants: int = 0
    partitions_dir: Path | None = None


class NationwideModelDatasetBuilder:
    """Build one common feature table from every qualified public plant partition."""

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
        self._reconciliation: dict[str, object] = {
            "revision_duplicate_rows": 0,
            "revision_conflict_keys": 0,
            "revision_selection_policy": "newest YYYYMMDD snapshot, then deterministic input order",
        }

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
        registry_path: Path | None = None
        quarantined_plants = 0
        if generation_paths is not None:
            paths = tuple(Path(path) for path in generation_paths)
            station_catalog = KmaStationCatalog.from_metadata(
                self.weather_root / "META_관측지점정보.csv"
            )
            registry_path = Path(destination).with_name("plant_registry.csv")
            registry = NationwidePlantRegistryBuilder(self.metadata, station_catalog).build(
                paths,
                registry_path,
                legacy_mapping=legacy_generation,
            )
            quarantined_plants = int(registry["model_ready_status"].eq("quarantined").sum())
            generation = self._from_standardized_generation(paths, registry)
        else:
            generation = legacy_generation

        years = sorted(generation["timestamp"].dt.year.unique())
        weather_paths = [self.weather_root / f"OBS_ASOS_TIM_{year}.csv" for year in years]
        missing_weather = [path for path in weather_paths if not path.exists()]
        if missing_weather:
            raise FileNotFoundError(f"KMA hourly files are missing: {missing_weather}")
        weather = KmaAsosNormalizer(self.weather_root / "META_관측지점정보.csv").read(
            weather_paths,
            station_ids=generation["station_id"].unique(),
        )
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
        context = ["timestamp", "company", "plant_id", "plant", "region"]
        for column in (
            "admin_province",
            "admin_city",
            "station_id",
            "weather_station_name",
            "weather_mapping_method",
            "weather_mapping_confidence",
            "weather_mapping_review_required",
            "energy_source",
        ):
            if column in result:
                context.append(column)
        result = result[[*context, *MODEL_READY_FEATURES, *QUALITY_COLUMNS, "generation_mwh"]]
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = destination.with_name(destination.name + ".tmp")
        compression = (
            {"method": "gzip", "compresslevel": 1, "mtime": 1}
            if destination.name.endswith(".gz")
            else None
        )
        result.to_csv(
            temporary,
            index=False,
            encoding="utf-8-sig",
            compression=compression,
        )
        temporary.replace(destination)
        partitions_dir = destination.parent / "model_ready_parts"
        partition_manifest = self._write_model_partitions(result, partitions_dir)

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
            "legacy_source_usage": (
                "reviewed KOSPO plant-to-ASOS mapping seed only; never target or time/plant filter"
            ),
            "plant_registry": str(registry_path) if registry_path else None,
            "plant_registry_policy": {
                "identity": "company + plant + energy_source",
                "administrative_region": "public address parsed separately from weather station",
                "weather_mapping": (
                    "reviewed legacy, <=50 km coordinate nearest, or unambiguous exact municipality"
                ),
                "ambiguous_mapping": "quarantine; never silently assign nearest city",
                "quarantined_plants": quarantined_plants,
            },
            "cross_partition_reconciliation": self._reconciliation,
            "dataset": str(destination),
            "partitioned_dataset": {
                "directory": str(partitions_dir),
                "partitioning": ["company", "year"],
                "files": partition_manifest,
            },
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
            "energy_source_filter": (
                "not applied in the nationwide model-ready table; each model job selects its technology"
            ),
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
            registry_path,
            quarantined_plants,
            partitions_dir,
        )

    def _from_standardized_generation(
        self,
        paths: Iterable[Path],
        registry: pd.DataFrame,
    ) -> pd.DataFrame:
        eligible_registry = registry.loc[
            registry["model_ready_status"].eq("eligible")
        ].copy()
        eligible_keys = set(
            eligible_registry[["company", "plant", "energy_source"]]
            .astype(str)
            .itertuples(index=False, name=None)
        )
        if not eligible_keys:
            raise ValueError("No plants have an auditable KMA station mapping")
        parts: list[pd.DataFrame] = []
        for partition_order, path in enumerate(paths):
            source = pd.read_csv(
                Path(path),
                usecols=[
                    "timestamp",
                    "company",
                    "plant_id",
                    "plant",
                    "energy_source",
                    "generation_mwh",
                ],
            )
            source["timestamp"] = pd.to_datetime(source["timestamp"], errors="coerce")
            keys = pd.MultiIndex.from_frame(
                source[["company", "plant", "energy_source"]].astype(str)
            )
            source = source.loc[keys.isin(eligible_keys)]
            if not source.empty:
                source["_partition_order"] = partition_order
                snapshot_dates = re.findall(r"20\d{6}", Path(path).name)
                source["_snapshot_date"] = max(map(int, snapshot_dates), default=0)
                parts.append(source)
        if not parts:
            raise ValueError("No standardized official generation rows match the registry quality gate")
        generation = pd.concat(parts, ignore_index=True)
        generation["generation_mwh"] = pd.to_numeric(generation["generation_mwh"], errors="coerce")
        generation = generation.dropna(subset=["timestamp", "generation_mwh"])
        identity = ["timestamp", "company", "plant_id"]
        duplicate_mask = generation.duplicated(identity, keep=False)
        conflicts = (
            generation.loc[duplicate_mask]
            .groupby(identity, sort=False)["generation_mwh"]
            .nunique(dropna=False)
            .gt(1)
        )
        self._reconciliation = {
            "revision_duplicate_rows": int(duplicate_mask.sum()),
            "revision_conflict_keys": int(conflicts.sum()),
            "revision_selection_policy": "newest YYYYMMDD snapshot, then deterministic input order",
        }
        # Public portals publish cumulative revisions. Prefer an explicit
        # YYYYMMDD snapshot suffix, then use deterministic input order as a
        # fallback. Summing two snapshots would double count generation.
        generation = generation.sort_values(
            ["_snapshot_date", "_partition_order"], kind="stable"
        ).drop_duplicates(
            identity,
            keep="last",
        )
        generation = generation.groupby(
            ["timestamp", "company", "plant", "energy_source"],
            as_index=False,
            sort=False,
        )["generation_mwh"].sum()

        mapping_columns = [
            "company",
            "plant",
            "energy_source",
            "capacity_mw",
            "tilt_deg",
            "admin_province",
            "admin_city",
            "weather_station_id",
            "weather_station_name",
            "weather_mapping_method",
            "weather_mapping_confidence",
            "weather_mapping_review_required",
        ]
        generation = generation.merge(
            eligible_registry[mapping_columns],
            on=["company", "plant", "energy_source"],
            how="inner",
            validate="many_to_one",
        )
        generation = generation.rename(columns={"weather_station_id": "station_id"})
        generation["region"] = generation["admin_province"].fillna(
            generation["admin_city"]
        ).fillna("unknown")
        generation["unit"] = ""
        generation["plant_id"] = generation["company"] + ":" + generation["plant"].astype(str)
        for column in ("latitude", "longitude", "address"):
            generation[column] = pd.NA
        generation["source_file"] = "official_standardized_partitions"
        if generation.duplicated(["timestamp", "plant_id"]).any():
            raise ValueError("Official aggregate contains duplicate timestamp/plant_id keys")
        return generation.sort_values(["timestamp", "plant_id"], kind="stable").reset_index(drop=True)

    @staticmethod
    def _write_model_partitions(
        frame: pd.DataFrame,
        destination: Path,
    ) -> list[dict[str, object]]:
        temporary_root = destination.with_name(destination.name + ".tmp")
        if temporary_root.exists():
            shutil.rmtree(temporary_root)
        temporary_root.mkdir(parents=True, exist_ok=False)
        working = frame.assign(_year=frame["timestamp"].dt.year)
        manifest: list[dict[str, object]] = []
        for (company, year), part in working.groupby(["company", "_year"], sort=True):
            relative = Path(f"company={company}") / f"year={int(year)}" / "part.csv.gz"
            output = temporary_root / relative
            output.parent.mkdir(parents=True, exist_ok=True)
            part.drop(columns="_year").to_csv(
                output,
                index=False,
                encoding="utf-8-sig",
                compression={"method": "gzip", "compresslevel": 1, "mtime": 1},
            )
            manifest.append(
                {
                    "path": str(destination / relative),
                    "company": str(company),
                    "year": int(year),
                    "rows": len(part),
                    "bytes": output.stat().st_size,
                }
            )
        previous = destination.with_name(destination.name + ".previous")
        if previous.exists():
            shutil.rmtree(previous)
        if destination.exists():
            shutil.move(str(destination), str(previous))
        shutil.move(str(temporary_root), str(destination))
        if previous.exists():
            shutil.rmtree(previous)
        return manifest


# Compatibility import for downstream notebooks. New code should use the
# nationwide name; the implementation no longer limits plants or dates.
LegacyModelDatasetBuilder = NationwideModelDatasetBuilder
