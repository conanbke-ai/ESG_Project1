from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from solar_forecast.artifacts.manifest import sha256_file, write_json_atomic

from .metadata import PlantMetadataCatalog
from .normalization import (
    ARCHIVE_WIDE_SCHEMAS,
    GENERATION_COLUMNS,
    GENERATION_CONTRACT_VERSION,
    DailyWideGenerationNormalizer,
    DailyWideSchema,
    KoenGenerationNormalizer,
    read_csv_with_fallback,
)


COMPANY_DIRECTORIES = {
    "koen": "한국남동발전",
    "kospo": "한국남부발전",
    "ewp": "한국동서발전",
    "iwest": "한국서부발전",
}


@dataclass(frozen=True)
class StandardizedPartition:
    company: str
    source: str
    source_bytes: int
    source_sha256: str
    destination: str
    destination_bytes: int
    destination_sha256: str
    rows: int
    plants: int
    start: str
    end: str
    capacity_coverage: float
    tilt_coverage: float
    coordinate_coverage: float
    duplicate_keys: int
    negative_generation: int
    energy_sources: dict[str, int]
    declared_source_unit: str | None
    resolved_source_unit: str | None
    unit_resolution_method: str | None
    unit_capacity_factor_p99: float | None
    unit_capacity_factor_max: float | None
    daily_total_resolution: dict[str, object] | None


@dataclass(frozen=True)
class StandardizationRun:
    output_dir: Path
    manifest_path: Path
    partitions: tuple[StandardizedPartition, ...]

    @property
    def rows(self) -> int:
        return sum(partition.rows for partition in self.partitions)


class GenerationSchemaRegistry:
    """Detect a known public CSV layout by its declared column contract."""

    def __init__(self, schemas: tuple[DailyWideSchema, ...] = ARCHIVE_WIDE_SCHEMAS):
        self.schemas = schemas

    @staticmethod
    def _required(schema: DailyWideSchema) -> set[str]:
        columns = {schema.date_column, *schema.hour_columns}
        if not schema.plant_from_filename:
            columns.add(schema.plant_column)
        for column in (
            schema.unit_column,
            schema.capacity_column,
            schema.tilt_column,
            schema.latitude_column,
            schema.longitude_column,
            schema.address_column,
            schema.energy_source_column,
            schema.include_column,
        ):
            if column:
                columns.add(column)
        columns.update(column for column in schema.id_columns if not column.startswith("_source_"))
        return columns

    def resolve(self, columns: set[str], company: str) -> DailyWideSchema:
        matches = [
            schema
            for schema in self.schemas
            if schema.company == company and self._required(schema).issubset(columns)
        ]
        if len(matches) != 1:
            raise ValueError(
                f"Expected one {company} schema match, found {len(matches)} for columns: {sorted(columns)}"
            )
        return matches[0]


class HistoricalGenerationStandardizationService:
    """Convert every retained generation export into partitioned common-schema data."""

    def __init__(
        self,
        input_root: Path,
        output_dir: Path,
        registry: GenerationSchemaRegistry | None = None,
        metadata: PlantMetadataCatalog | None = None,
    ):
        self.input_root = Path(input_root)
        self.output_dir = Path(output_dir)
        self.registry = registry or GenerationSchemaRegistry()
        self.metadata = metadata or PlantMetadataCatalog.from_directory(self.input_root / "location")

    def run(self) -> StandardizationRun:
        partitions: list[StandardizedPartition] = []
        errors: list[dict[str, str]] = []
        generation_dir = self.output_dir / "generation"
        generation_dir.mkdir(parents=True, exist_ok=True)
        for company, directory_name in COMPANY_DIRECTORIES.items():
            source_dir = self.input_root / directory_name
            for source in sorted(source_dir.rglob("*.csv")):
                try:
                    normalized = self._normalize(company, source)
                    normalization_attrs = normalized.attrs.copy()
                    normalized = self.metadata.enrich(normalized)
                    normalized.attrs.update(normalization_attrs)
                    destination = generation_dir / company / f"{source.stem}_standardized.csv.gz"
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    temporary = destination.with_name(destination.name + ".tmp")
                    normalized.to_csv(
                        temporary,
                        index=False,
                        encoding="utf-8-sig",
                        compression={"method": "gzip", "compresslevel": 1, "mtime": 1},
                    )
                    temporary.replace(destination)
                    partitions.append(self._audit(company, source, destination, normalized))
                except Exception as exc:
                    errors.append({"company": company, "source": str(source), "error": str(exc)})
        manifest = {
            "created_at": datetime.now().isoformat(),
            "contract_version": GENERATION_CONTRACT_VERSION,
            "contract": GENERATION_COLUMNS,
            "input_root": str(self.input_root),
            "processing_model": "one source file in memory + deterministic gzip partition",
            "partitions": [asdict(partition) for partition in partitions],
            "errors": errors,
            "summary": {
                "files": len(partitions),
                "rows": sum(partition.rows for partition in partitions),
                "companies": sorted({partition.company for partition in partitions}),
                "input_bytes": sum(partition.source_bytes for partition in partitions),
                "output_bytes": sum(partition.destination_bytes for partition in partitions),
            },
        }
        manifest_path = self.output_dir / "generation_manifest.json"
        write_json_atomic(manifest_path, manifest)
        if errors:
            samples = "; ".join(f"{Path(item['source']).name}: {item['error']}" for item in errors[:3])
            raise ValueError(f"Failed to standardize {len(errors)} files. {samples}")
        return StandardizationRun(self.output_dir, manifest_path, tuple(partitions))

    def _normalize(self, company: str, source: Path) -> pd.DataFrame:
        if company == "koen":
            return KoenGenerationNormalizer().read(source)
        raw = read_csv_with_fallback(source)
        raw.columns = [str(column).strip() for column in raw.columns]
        schema = self.registry.resolve(set(raw.columns), company)
        normalizer = DailyWideGenerationNormalizer(schema)
        if schema.plant_from_filename:
            return normalizer.read(source)
        return normalizer.transform(raw, source_file=source.name)

    @staticmethod
    def _audit(
        company: str,
        source: Path,
        destination: Path,
        frame: pd.DataFrame,
    ) -> StandardizedPartition:
        duplicate_keys = int(frame.duplicated(["timestamp", "plant_id"]).sum())
        if duplicate_keys:
            raise ValueError(f"Cross-row duplicate keys remain after normalization: {duplicate_keys}")
        unit_resolution = frame.attrs.get("generation_unit_resolution", {})
        return StandardizedPartition(
            company=company,
            source=str(source),
            source_bytes=source.stat().st_size,
            source_sha256=sha256_file(source),
            destination=str(destination),
            destination_bytes=destination.stat().st_size,
            destination_sha256=sha256_file(destination),
            rows=len(frame),
            plants=frame["plant_id"].nunique(),
            start=frame["timestamp"].min().isoformat(),
            end=frame["timestamp"].max().isoformat(),
            capacity_coverage=float(frame["capacity_mw"].notna().mean()),
            tilt_coverage=float(frame["tilt_deg"].notna().mean()),
            coordinate_coverage=float(frame[["latitude", "longitude"]].notna().all(axis=1).mean()),
            duplicate_keys=duplicate_keys,
            negative_generation=int(frame["generation_mwh"].lt(0).sum()),
            energy_sources={
                str(source): int(count)
                for source, count in frame["energy_source"].value_counts(dropna=False).items()
            },
            declared_source_unit=unit_resolution.get("declared_source_unit"),
            resolved_source_unit=unit_resolution.get("resolved_source_unit"),
            unit_resolution_method=unit_resolution.get("method"),
            unit_capacity_factor_p99=unit_resolution.get("capacity_factor_p99"),
            unit_capacity_factor_max=unit_resolution.get("capacity_factor_max"),
            daily_total_resolution=unit_resolution.get("daily_total"),
        )
