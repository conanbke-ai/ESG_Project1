"""Memory-bounded aggregation of the KPX EPSIS national solar inventory.

The EPSIS export is a generator-registration inventory, not a training-data
catalogue and not a list of unique physical plants.  This module deliberately
keeps those semantics: exact duplicate rows are measured but retained in every
count and capacity total.
"""

from __future__ import annotations

from collections import Counter
from contextlib import contextmanager
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from hashlib import sha256
import csv
import json
from pathlib import Path
import re
import sqlite3
import tempfile
from typing import Any, Iterator, Mapping, Sequence
import unicodedata


REQUIRED_COLUMNS: tuple[str, ...] = (
    "회사명",
    "발전기명",
    "호기",
    "설비용량",
    "회원구분",
    "급전방식",
    "발전원",
    "발전종류",
    "사업구분",
    "광역지역",
    "세부지역",
)

CANONICAL_REGIONS: tuple[str, ...] = (
    "서울특별시",
    "부산광역시",
    "대구광역시",
    "인천광역시",
    "광주광역시",
    "대전광역시",
    "울산광역시",
    "세종특별자치시",
    "경기도",
    "강원특별자치도",
    "충청북도",
    "충청남도",
    "전북특별자치도",
    "전라남도",
    "경상북도",
    "경상남도",
    "제주특별자치도",
)

REGION_ALIASES: dict[str, str] = {
    "서울": "서울특별시",
    "서울시": "서울특별시",
    "서울특별시": "서울특별시",
    "부산": "부산광역시",
    "부산시": "부산광역시",
    "부산광역시": "부산광역시",
    "대구": "대구광역시",
    "대구시": "대구광역시",
    "대구광역시": "대구광역시",
    "인천": "인천광역시",
    "인천시": "인천광역시",
    "인천광역시": "인천광역시",
    "광주": "광주광역시",
    "광주시": "광주광역시",
    "광주광역시": "광주광역시",
    "대전": "대전광역시",
    "대전시": "대전광역시",
    "대전광역시": "대전광역시",
    "울산": "울산광역시",
    "울산시": "울산광역시",
    "울산광역시": "울산광역시",
    "세종": "세종특별자치시",
    "세종시": "세종특별자치시",
    "세종특별자치시": "세종특별자치시",
    "경기": "경기도",
    "경기도": "경기도",
    "강원": "강원특별자치도",
    "강원도": "강원특별자치도",
    "강원특별자치도": "강원특별자치도",
    "충북": "충청북도",
    "충청북도": "충청북도",
    "충남": "충청남도",
    "충청남도": "충청남도",
    "전북": "전북특별자치도",
    "전라북도": "전북특별자치도",
    "전북특별자치도": "전북특별자치도",
    "전남": "전라남도",
    "전라남도": "전라남도",
    "경북": "경상북도",
    "경상북도": "경상북도",
    "경남": "경상남도",
    "경상남도": "경상남도",
    "제주": "제주특별자치도",
    "제주도": "제주특별자치도",
    "제주특별자치도": "제주특별자치도",
}

# Administrative-centre coordinates are stable, documented fallbacks only.
# A matched subregion coordinate always takes precedence.
PROVINCE_CENTROIDS: dict[str, tuple[float, float]] = {
    "서울특별시": (37.5665, 126.9780),
    "부산광역시": (35.1796, 129.0756),
    "대구광역시": (35.8714, 128.6014),
    "인천광역시": (37.4563, 126.7052),
    "광주광역시": (35.1595, 126.8526),
    "대전광역시": (36.3504, 127.3845),
    "울산광역시": (35.5384, 129.3114),
    "세종특별자치시": (36.4800, 127.2890),
    "경기도": (37.4138, 127.5183),
    "강원특별자치도": (37.8228, 128.1555),
    "충청북도": (36.6357, 127.4917),
    "충청남도": (36.6588, 126.6728),
    "전북특별자치도": (35.8203, 127.1088),
    "전라남도": (34.8161, 126.4629),
    "경상북도": (36.4919, 128.8889),
    "경상남도": (35.4606, 128.2132),
    "제주특별자치도": (33.4996, 126.5312),
}

FOOTER_CAPACITY_TOLERANCE_MW = Decimal("0.000001")


class InventoryConfigurationError(ValueError):
    """Raised when source configuration is incomplete or inconsistent."""


class InventorySchemaError(ValueError):
    """Raised when the EPSIS CSV no longer satisfies its column contract."""


class InventoryIntegrityError(ValueError):
    """Raised when a local artifact does not match its configured digest."""


class InventoryContentError(ValueError):
    """Raised when rows violate the configured solar-inventory contract."""


@dataclass(frozen=True)
class NationalInventoryConfig:
    """Resolved source metadata used by the repository and output lineage."""

    project_root: Path
    local_path: Path
    local_path_label: str
    coordinate_cache_path: Path
    boundary_path: Path
    boundary_path_label: str
    expected_sha256: str
    provider: str
    source_url: str
    reference_date: str
    downloaded_at: str
    encoding: str
    scope: str
    limitations: tuple[str, ...]
    dataset_id: str = "kpx_epsis_national_solar_generator_status"
    source_system: str = "EPSIS"
    capacity_unit: str = "MW"
    record_unit: str = "generator_registration_row"

    @classmethod
    def from_mapping(
        cls,
        values: Mapping[str, Any],
        *,
        project_root: str | Path = ".",
    ) -> "NationalInventoryConfig":
        required = (
            "local_path",
            "expected_sha256",
            "provider",
            "source_url",
            "reference_date",
            "downloaded_at",
            "encoding",
            "scope",
            "limitations",
        )
        missing = [key for key in required if key not in values]
        if missing:
            raise InventoryConfigurationError(
                "Missing national inventory config keys: " + ", ".join(missing)
            )

        root = Path(project_root).resolve()
        local_label = str(values["local_path"])
        local_path = Path(local_label)
        if not local_path.is_absolute():
            local_path = root / local_path
        coordinate_label = str(
            values.get("coordinate_cache_path", "map/json/coord_cache.json")
        )
        coordinate_path = Path(coordinate_label)
        if not coordinate_path.is_absolute():
            coordinate_path = root / coordinate_path
        boundary_label = str(values.get("boundary_path", "map/json/geoJson.json"))
        boundary_path = Path(boundary_label)
        if not boundary_path.is_absolute():
            boundary_path = root / boundary_path

        limitations_value = values["limitations"]
        if isinstance(limitations_value, str) or not isinstance(
            limitations_value, Sequence
        ):
            raise InventoryConfigurationError("limitations must be a list of strings")
        limitations = tuple(str(item) for item in limitations_value)

        expected_digest = str(values["expected_sha256"]).strip().lower()
        if not re.fullmatch(r"[0-9a-f]{64}", expected_digest):
            raise InventoryConfigurationError("expected_sha256 must be 64 hexadecimal characters")

        return cls(
            project_root=root,
            local_path=local_path.resolve(),
            local_path_label=local_label.replace("\\", "/"),
            coordinate_cache_path=coordinate_path.resolve(),
            boundary_path=boundary_path.resolve(),
            boundary_path_label=boundary_label.replace("\\", "/"),
            expected_sha256=expected_digest,
            provider=str(values["provider"]),
            source_url=str(values["source_url"]),
            reference_date=str(values["reference_date"]),
            downloaded_at=str(values["downloaded_at"]),
            encoding=str(values["encoding"]),
            scope=str(values["scope"]),
            limitations=limitations,
            dataset_id=str(values.get("dataset_id", cls.dataset_id)),
            source_system=str(values.get("source_system", cls.source_system)),
            capacity_unit=str(values.get("capacity_unit", cls.capacity_unit)),
            record_unit=str(values.get("record_unit", cls.record_unit)),
        )

    @classmethod
    def from_json(
        cls,
        path: str | Path,
        *,
        project_root: str | Path | None = None,
    ) -> "NationalInventoryConfig":
        config_path = Path(path).resolve()
        values = json.loads(config_path.read_text(encoding="utf-8"))
        if not isinstance(values, dict):
            raise InventoryConfigurationError("national inventory config must be a JSON object")
        if project_root is None:
            project_root = config_path.parent.parent
        return cls.from_mapping(values, project_root=project_root)


@dataclass(frozen=True)
class InventoryCsvRecord:
    values: Mapping[str, str]
    raw_values: tuple[str, ...]
    malformed: bool
    row_number: int


class KpxEpsisSolarInventoryRepository:
    """Verify and stream an EPSIS CSV without materialising it as a DataFrame."""

    def __init__(self, config: NationalInventoryConfig):
        self.config = config

    def verify_sha256(self) -> str:
        if not self.config.local_path.is_file():
            raise FileNotFoundError(self.config.local_path)
        digest = sha256()
        with self.config.local_path.open("rb") as source:
            while chunk := source.read(1024 * 1024):
                digest.update(chunk)
        actual = digest.hexdigest()
        if actual != self.config.expected_sha256:
            raise InventoryIntegrityError(
                "National inventory SHA-256 mismatch: "
                f"expected={self.config.expected_sha256}, actual={actual}"
            )
        return actual

    @contextmanager
    def stream(
        self,
    ) -> Iterator[tuple[tuple[str, ...], Iterator[InventoryCsvRecord]]]:
        """Yield validated columns and a lazy row iterator while the file is open."""

        with self.config.local_path.open(
            "r", encoding=self.config.encoding, errors="strict", newline=""
        ) as source:
            reader = csv.reader(source)
            try:
                raw_header = next(reader)
            except StopIteration as error:
                raise InventorySchemaError("National inventory CSV is empty") from error
            header = tuple(
                value.lstrip("\ufeff").strip() if index == 0 else value.strip()
                for index, value in enumerate(raw_header)
            )
            if len(set(header)) != len(header):
                raise InventorySchemaError("National inventory CSV contains duplicate columns")
            missing = [column for column in REQUIRED_COLUMNS if column not in header]
            if missing:
                raise InventorySchemaError(
                    "Missing required EPSIS columns: " + ", ".join(missing)
                )

            def records() -> Iterator[InventoryCsvRecord]:
                for row_number, row in enumerate(reader, start=2):
                    malformed = len(row) != len(header)
                    padded = row[: len(header)] + [""] * max(0, len(header) - len(row))
                    yield InventoryCsvRecord(
                        values=dict(zip(header, padded)),
                        raw_values=tuple(row),
                        malformed=malformed,
                        row_number=row_number,
                    )

            yield header, records()


@dataclass
class _Aggregate:
    records: int = 0
    capacity_mw: Decimal = Decimal("0")

    def add(self, capacity_mw: Decimal | None) -> None:
        self.records += 1
        if capacity_mw is not None:
            self.capacity_mw += capacity_mw


@dataclass
class _LocationAggregate(_Aggregate):
    source_locations: set[str] = field(default_factory=set)


class _DiskDuplicateTracker:
    """Exact row tracking backed by temporary SQLite, keeping RAM bounded."""

    def __enter__(self) -> "_DiskDuplicateTracker":
        self._temporary = tempfile.TemporaryDirectory(prefix="kpx-inventory-")
        database = Path(self._temporary.name) / "rows.sqlite3"
        self._connection = sqlite3.connect(database)
        self._connection.execute("PRAGMA journal_mode=OFF")
        self._connection.execute("PRAGMA synchronous=OFF")
        self._connection.execute(
            "CREATE TABLE seen_rows (payload TEXT PRIMARY KEY) WITHOUT ROWID"
        )
        return self

    def is_duplicate(self, values: tuple[str, ...]) -> bool:
        payload = json.dumps(values, ensure_ascii=False, separators=(",", ":"))
        cursor = self._connection.execute(
            "INSERT OR IGNORE INTO seen_rows(payload) VALUES (?)", (payload,)
        )
        return cursor.rowcount == 0

    def __exit__(self, exc_type: Any, exc: Any, traceback: Any) -> None:
        self._connection.close()
        self._temporary.cleanup()


class NationalInventoryService:
    """Aggregate national/regional totals and attach auditable map coordinates."""

    def __init__(
        self,
        repository: KpxEpsisSolarInventoryRepository,
        *,
        coordinate_cache: Mapping[str, Sequence[float]] | None = None,
    ):
        self.repository = repository
        self._coordinate_cache = (
            dict(coordinate_cache) if coordinate_cache is not None else None
        )

    @classmethod
    def from_config(
        cls,
        config: Mapping[str, Any] | str | Path,
        *,
        project_root: str | Path = ".",
        coordinate_cache: Mapping[str, Sequence[float]] | None = None,
    ) -> "NationalInventoryService":
        if isinstance(config, Mapping):
            resolved = NationalInventoryConfig.from_mapping(
                config, project_root=project_root
            )
        else:
            resolved = NationalInventoryConfig.from_json(
                config, project_root=project_root
            )
        return cls(
            KpxEpsisSolarInventoryRepository(resolved),
            coordinate_cache=coordinate_cache,
        )

    def build(self) -> dict[str, Any]:
        digest = self.repository.verify_sha256()
        region_totals = {region: _Aggregate() for region in CANONICAL_REGIONS}
        location_totals: dict[tuple[str, str], _LocationAggregate] = {}
        missing_cells: Counter[str] = Counter()
        coordinate_cache = self._load_coordinate_cache()

        total = _Aggregate()
        duplicate_records = 0
        duplicate_capacity = Decimal("0")
        negative_capacity = 0
        zero_capacity = 0
        invalid_capacity = 0
        malformed_records = 0
        replacement_cells = 0
        fullwidth_question_cells = 0
        garbled_records = 0
        unknown_region_records = 0
        footer: InventoryCsvRecord | None = None

        with _DiskDuplicateTracker() as duplicates:
            with self.repository.stream() as (columns, records):
                pending: InventoryCsvRecord | None = None
                for record in records:
                    if pending is not None:
                        metrics = self._accumulate_record(
                            pending,
                            total=total,
                            region_totals=region_totals,
                            location_totals=location_totals,
                            missing_cells=missing_cells,
                            duplicates=duplicates,
                        )
                        duplicate_records += metrics["duplicate_records"]
                        duplicate_capacity += metrics["duplicate_capacity_mw"]
                        negative_capacity += metrics["negative_capacity_records"]
                        zero_capacity += metrics["zero_capacity_records"]
                        invalid_capacity += metrics["invalid_capacity_records"]
                        malformed_records += metrics["malformed_records"]
                        replacement_cells += metrics["replacement_character_cells"]
                        fullwidth_question_cells += metrics[
                            "fullwidth_question_mark_cells"
                        ]
                        garbled_records += metrics["garbled_records"]
                        unknown_region_records += metrics["unknown_region_records"]
                    pending = record

                if pending is not None and self._is_footer(pending):
                    footer = pending
                elif pending is not None:
                    metrics = self._accumulate_record(
                        pending,
                        total=total,
                        region_totals=region_totals,
                        location_totals=location_totals,
                        missing_cells=missing_cells,
                        duplicates=duplicates,
                    )
                    duplicate_records += metrics["duplicate_records"]
                    duplicate_capacity += metrics["duplicate_capacity_mw"]
                    negative_capacity += metrics["negative_capacity_records"]
                    zero_capacity += metrics["zero_capacity_records"]
                    invalid_capacity += metrics["invalid_capacity_records"]
                    malformed_records += metrics["malformed_records"]
                    replacement_cells += metrics["replacement_character_cells"]
                    fullwidth_question_cells += metrics[
                        "fullwidth_question_mark_cells"
                    ]
                    garbled_records += metrics["garbled_records"]
                    unknown_region_records += metrics["unknown_region_records"]

        regions = self._region_records(region_totals)
        locations, coordinate_counts, invalid_cache_coordinates = self._location_records(
            location_totals, coordinate_cache
        )
        footer_quality = self._footer_quality(footer, total)

        config = self.repository.config
        inventory = {
            "source": {
                "dataset_id": config.dataset_id,
                "provider": config.provider,
                "source_system": config.source_system,
                "source_url": config.source_url,
                "local_path": config.local_path_label,
                "reference_date": config.reference_date,
                "downloaded_at": config.downloaded_at,
                "encoding": config.encoding,
                "capacity_unit": config.capacity_unit,
                "record_unit": config.record_unit,
                "scope": config.scope,
                "limitations": list(config.limitations),
                "boundary_path": config.boundary_path_label,
                "sha256": digest,
                "expected_sha256": config.expected_sha256,
                "sha256_verified": True,
                "bytes": config.local_path.stat().st_size,
            },
            "summary": {
                "generator_records": total.records,
                "total_capacity_mw": _number(total.capacity_mw),
                "canonical_regions": len(CANONICAL_REGIONS),
                "regions_with_records": sum(
                    aggregate.records > 0 for aggregate in region_totals.values()
                ),
                "source_subregion_labels": len(
                    {
                        location
                        for aggregate in location_totals.values()
                        for location in aggregate.source_locations
                    }
                ),
                "subregions": len(location_totals),
                "located_subregions": sum(
                    count
                    for basis, count in coordinate_counts.items()
                    if basis != "unresolved"
                ),
            },
            "regions": regions,
            "locations": locations,
            "quality": {
                "schema_valid": True,
                "required_columns": list(REQUIRED_COLUMNS),
                "source_columns": list(columns),
                "malformed_records": malformed_records,
                "invalid_capacity_records": invalid_capacity,
                "negative_capacity_records": negative_capacity,
                "zero_capacity_records": zero_capacity,
                "exact_duplicate_records": duplicate_records,
                "exact_duplicate_capacity_mw": _number(duplicate_capacity),
                "replacement_character_cells": replacement_cells,
                "fullwidth_question_mark_cells": fullwidth_question_cells,
                "garbled_records": garbled_records,
                "unknown_region_records": unknown_region_records,
                "missing_cells_by_column": dict(sorted(missing_cells.items())),
                "coordinate_basis_counts": dict(sorted(coordinate_counts.items())),
                "invalid_coordinate_cache_entries": invalid_cache_coordinates,
                "duplicates_retained": True,
                "footer": footer_quality,
            },
        }
        return {"national_inventory": inventory}

    def _accumulate_record(
        self,
        record: InventoryCsvRecord,
        *,
        total: _Aggregate,
        region_totals: dict[str, _Aggregate],
        location_totals: dict[tuple[str, str], _LocationAggregate],
        missing_cells: Counter[str],
        duplicates: _DiskDuplicateTracker,
    ) -> dict[str, Any]:
        values = record.values
        energy_source = _clean(values.get("발전원", ""))
        if energy_source != "태양에너지":
            raise InventoryContentError(
                "National solar inventory contains a non-solar row: "
                f"CSV row {record.row_number}, 발전원={energy_source or '<empty>'}"
            )
        capacity = _decimal(values.get("설비용량", ""))
        duplicate = duplicates.is_duplicate(record.raw_values)
        region_raw = _clean(values.get("광역지역", ""))
        region = canonical_region(region_raw)
        unknown_region = region not in CANONICAL_REGIONS
        if unknown_region and region not in region_totals:
            region_totals[region] = _Aggregate()
        source_subregion = _clean(values.get("세부지역", ""))
        subregion = canonical_location(source_subregion, region)
        if not subregion:
            subregion = region

        total.add(capacity)
        region_totals[region].add(capacity)
        location = location_totals.setdefault(
            (region, subregion), _LocationAggregate()
        )
        location.add(capacity)
        if source_subregion:
            location.source_locations.add(source_subregion)

        for column in REQUIRED_COLUMNS:
            if not _clean(values.get(column, "")):
                missing_cells[column] += 1

        replacement = sum("\ufffd" in value for value in record.raw_values)
        fullwidth_question = sum("\uff1f" in value for value in record.raw_values)
        return {
            "duplicate_records": int(duplicate),
            "duplicate_capacity_mw": capacity if duplicate and capacity is not None else Decimal("0"),
            "negative_capacity_records": int(capacity is not None and capacity < 0),
            "zero_capacity_records": int(capacity is not None and capacity == 0),
            "invalid_capacity_records": int(capacity is None),
            "malformed_records": int(record.malformed),
            "replacement_character_cells": replacement,
            "fullwidth_question_mark_cells": fullwidth_question,
            "garbled_records": int(bool(replacement or fullwidth_question)),
            "unknown_region_records": int(unknown_region),
        }

    def _load_coordinate_cache(self) -> dict[str, Sequence[float]]:
        if self._coordinate_cache is not None:
            return self._coordinate_cache
        path = self.repository.config.coordinate_cache_path
        if not path.exists():
            return {}
        payload = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            raise InventoryConfigurationError("coordinate cache must be a JSON object")
        return payload

    @staticmethod
    def _is_footer(record: InventoryCsvRecord) -> bool:
        return (
            _clean(record.values.get("발전기명", "")) == "통합"
            and not _clean(record.values.get("광역지역", ""))
            and not _clean(record.values.get("세부지역", ""))
        )

    @staticmethod
    def _footer_quality(
        footer: InventoryCsvRecord | None, total: _Aggregate
    ) -> dict[str, Any]:
        if footer is None:
            return {
                "found": False,
                "excluded_records": 0,
                "declared_record_count": None,
                "declared_capacity_mw": None,
                "record_count_matches": False,
                "capacity_matches": False,
                "capacity_difference_mw": None,
            }
        declared_count = _integer(footer.values.get("호기", ""))
        declared_capacity = _decimal(footer.values.get("설비용량", ""))
        difference = (
            total.capacity_mw - declared_capacity
            if declared_capacity is not None
            else None
        )
        return {
            "found": True,
            "excluded_records": 1,
            "declared_record_count": declared_count,
            "declared_capacity_mw": _number(declared_capacity),
            "record_count_matches": declared_count == total.records,
            "capacity_matches": (
                difference is not None
                and abs(difference) <= FOOTER_CAPACITY_TOLERANCE_MW
            ),
            "capacity_difference_mw": _number(difference),
        }

    @staticmethod
    def _region_records(
        totals: Mapping[str, _Aggregate],
    ) -> list[dict[str, Any]]:
        ordered = list(CANONICAL_REGIONS) + sorted(
            region for region in totals if region not in CANONICAL_REGIONS
        )
        return [
            {
                "region": region,
                "generator_records": totals[region].records,
                "capacity_mw": _number(totals[region].capacity_mw),
            }
            for region in ordered
        ]

    @staticmethod
    def _location_records(
        totals: Mapping[tuple[str, str], _LocationAggregate],
        cache: Mapping[str, Sequence[float]],
    ) -> tuple[list[dict[str, Any]], Counter[str], int]:
        exact: dict[str, tuple[float, float]] = {}
        normalized: dict[str, tuple[str, tuple[float, float]]] = {}
        invalid_entries = 0
        for key, value in cache.items():
            coordinates = _coordinates(value)
            if coordinates is None:
                invalid_entries += 1
                continue
            clean_key = _clean(key)
            exact[clean_key] = coordinates
            normalized.setdefault(
                normalize_location_key(clean_key), (clean_key, coordinates)
            )

        records: list[dict[str, Any]] = []
        bases: Counter[str] = Counter()
        for (region, subregion), aggregate in sorted(totals.items()):
            exact_key = next(
                (key for key in sorted(aggregate.source_locations) if key in exact),
                None,
            )
            if exact_key is not None:
                latitude, longitude = exact[exact_key]
                basis = "exact"
                coordinate_key: str | None = exact_key
            else:
                match = next(
                    (
                        normalized[normalize_location_key(key, region)]
                        for key in sorted(aggregate.source_locations)
                        if normalize_location_key(key, region) in normalized
                    ),
                    None,
                )
                if match is not None:
                    coordinate_key, (latitude, longitude) = match
                    basis = "normalized"
                elif region in PROVINCE_CENTROIDS:
                    latitude, longitude = PROVINCE_CENTROIDS[region]
                    basis = "province_centroid"
                    coordinate_key = region
                else:
                    latitude = longitude = None
                    basis = "unresolved"
                    coordinate_key = None
            bases[basis] += 1
            records.append(
                {
                    "region": region,
                    "subregion": subregion,
                    "generator_records": aggregate.records,
                    "capacity_mw": _number(aggregate.capacity_mw),
                    "latitude": latitude,
                    "longitude": longitude,
                    "coordinate_basis": basis,
                    "coordinate_key": coordinate_key,
                }
            )
        return records, bases, invalid_entries


def build_national_inventory(
    project_root: str | Path,
    *,
    config_path: str | Path = "config/national_solar_inventory.json",
) -> dict[str, Any]:
    """Convenience entry point for dashboard builders and command-line jobs."""

    root = Path(project_root).resolve()
    path = Path(config_path)
    if not path.is_absolute():
        path = root / path
    service = NationalInventoryService.from_config(path, project_root=root)
    return service.build()


def canonical_region(value: object) -> str:
    cleaned = _clean(value)
    return REGION_ALIASES.get(cleaned, cleaned or "미분류")


def canonical_location(value: object, region: str) -> str:
    text = _clean(value)
    if not text:
        return ""
    for alias in sorted(REGION_ALIASES, key=len, reverse=True):
        if text == alias or text.startswith(alias + " "):
            suffix = text[len(alias) :].strip()
            canonical = REGION_ALIASES[alias]
            return f"{canonical} {suffix}".strip()
    if region in CANONICAL_REGIONS:
        return f"{region} {text}"
    return text


def normalize_location_key(value: object, region: str | None = None) -> str:
    text = canonical_location(
        value, region if region is not None else canonical_region(_first_token(value))
    )
    return re.sub(r"\s+", "", text)


def _first_token(value: object) -> str:
    text = _clean(value)
    return text.split(" ", 1)[0] if text else ""


def _clean(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(value))).strip()


def _decimal(value: object) -> Decimal | None:
    text = _clean(value).replace(",", "")
    if not text:
        return None
    try:
        result = Decimal(text)
    except InvalidOperation:
        return None
    return result if result.is_finite() else None


def _integer(value: object) -> int | None:
    number = _decimal(value)
    if number is None or number != number.to_integral_value():
        return None
    return int(number)


def _number(value: Decimal | None) -> float | None:
    return float(value) if value is not None else None


def _coordinates(value: Sequence[float] | object) -> tuple[float, float] | None:
    if isinstance(value, (str, bytes)) or not isinstance(value, Sequence):
        return None
    if len(value) != 2:
        return None
    try:
        latitude, longitude = float(value[0]), float(value[1])
    except (TypeError, ValueError):
        return None
    if not (32.0 <= latitude <= 39.5 and 124.0 <= longitude <= 132.0):
        return None
    return latitude, longitude


__all__ = [
    "CANONICAL_REGIONS",
    "KpxEpsisSolarInventoryRepository",
    "NationalInventoryConfig",
    "NationalInventoryService",
    "REQUIRED_COLUMNS",
    "InventoryConfigurationError",
    "InventoryContentError",
    "InventoryIntegrityError",
    "InventorySchemaError",
    "build_national_inventory",
    "canonical_region",
]
