from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
from typing import Any

import pandas as pd

from solar_forecast.collectors.csv_artifacts import inspect_csv_artifact
from solar_forecast.collectors.naming import is_canonical_solar_download_name
from solar_forecast.collectors.normalization import read_csv_with_fallback
from solar_forecast.reporting.national_inventory import build_national_inventory


COMPANY_NAMES = {
    "ewp": "한국동서발전(주)",
    "kospo": "한국남부발전(주)",
    "iwest": "한국서부발전(주)",
    "koen": "한국남동발전(주)",
    "krc": "한국농어촌공사",
    "komipo": "한국중부발전(주)",
}


@dataclass(frozen=True)
class DashboardBuildResult:
    data_path: Path
    boundary_path: Path | None
    solar_dashboard: Path
    mapping_report: Path
    national_generator_records: int
    national_capacity_mw: float
    solar_assets: int
    eligible_solar_assets: int


class DashboardBuilder:
    """Build the local, data-backed dashboard payload from current artifacts."""

    def __init__(self, project_root: Path, output_dir: Path | None = None):
        self.project_root = project_root.resolve()
        self.output_dir = (output_dir or self.project_root / "dashboard").resolve()

    def build(self) -> DashboardBuildResult:
        registry_path = self.project_root / "file/standardized/plant_registry.csv"
        quality_path = self.project_root / "file/standardized/plant_quality_report.csv"
        model_manifest_path = self.project_root / "file/standardized/model_ready_manifest.json"
        if not registry_path.exists():
            raise FileNotFoundError(f"Run prepare-data first: {registry_path}")

        registry = read_csv_with_fallback(registry_path)
        quality = (
            read_csv_with_fallback(quality_path)
            if quality_path.exists()
            else pd.DataFrame(columns=["plant_id"])
        )
        manifest = (
            json.loads(model_manifest_path.read_text(encoding="utf-8"))
            if model_manifest_path.exists()
            else {}
        )
        stations = self._load_weather_stations()
        national = build_national_inventory(self.project_root)["national_inventory"]
        payload = self._payload(registry, quality, manifest, stations, national)
        self._publish_static_assets()

        data_path = self.output_dir / "data/dashboard_data.json"
        data_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = data_path.with_suffix(".json.part")
        temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False),
            encoding="utf-8",
        )
        temporary.replace(data_path)
        boundary_path = self._copy_province_boundaries(national)

        solar = registry.loc[registry["energy_source"].eq("solar")]
        eligible = solar["model_ready_status"].eq("eligible")
        return DashboardBuildResult(
            data_path=data_path,
            boundary_path=boundary_path,
            solar_dashboard=self.output_dir / "solar_dashboard.html",
            mapping_report=self.output_dir / "plant_region_report_perm.html",
            national_generator_records=int(national["summary"]["generator_records"]),
            national_capacity_mw=float(national["summary"]["total_capacity_mw"]),
            solar_assets=int(len(solar)),
            eligible_solar_assets=int(eligible.sum()),
        )

    def _payload(
        self,
        registry: pd.DataFrame,
        quality: pd.DataFrame,
        manifest: dict[str, Any],
        stations: dict[int, dict[str, Any]],
        national_inventory: dict[str, Any],
    ) -> dict[str, Any]:
        solar = registry.loc[registry["energy_source"].eq("solar")].copy()
        quality_columns = [
            "plant_id",
            "rows",
            "hourly_coverage",
            "missing_generation_rate",
            "negative_generation_rate",
            "capacity_exceeded_rate",
            "daylight_zero_rate",
            "flatline_rate",
            "missing_weather_rate",
            "sensor_risk",
            "recommended_action",
        ]
        available = [column for column in quality_columns if column in quality]
        solar = solar.merge(quality[available], on="plant_id", how="left")

        plants = [self._plant_record(row, stations) for _, row in solar.iterrows()]
        companies = self._group_summary(solar, "company", company_labels=True)
        regions = self._group_summary(solar, "admin_province")
        risk_counts = Counter(
            str(value) for value in solar.get("sensor_risk", pd.Series(dtype=str)).dropna()
        )
        validation = self._validate_registry(registry)
        gold_consistency = self._audit_gold_partitions(registry)
        if gold_consistency["available"]:
            validation.append(
                {
                    "check": "gold_partition_registry_consistency",
                    "passed": gold_consistency["violations"] == 0,
                    "violations": gold_consistency["violations"],
                }
            )

        return {
            "meta": {
                "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
                "scope": "전국 발전설비 현황과 발전량 학습 포트폴리오를 분리한 로컬 대시보드",
                "training_scope": "공개 시간별 발전실적과 기상자료를 함께 확보한 태양광 학습 포트폴리오",
                "location_policy": (
                    "발전소 실좌표가 없으면 학습에 사용한 ASOS 관측소 좌표를 대리표시하며 "
                    "두 좌표 유형을 구분한다"
                ),
                "raw_policy": "Bronze 원본 바이트 보존; Silver/Gold만 UTF-8-SIG와 표준 컬럼으로 변환",
            },
            "national_inventory": national_inventory,
            "summary": {
                "registered_assets_all_energy": int(len(registry)),
                "solar_assets": int(len(solar)),
                "eligible_solar_assets": int(solar["model_ready_status"].eq("eligible").sum()),
                "quarantined_solar_assets": int(solar["model_ready_status"].eq("quarantined").sum()),
                "companies": int(solar["company"].nunique()),
                "model_rows_all_energy": int(manifest.get("rows", 0)),
                "model_plants_all_energy": int(manifest.get("plants", 0)),
                "generation_start": manifest.get("start"),
                "generation_end": manifest.get("end"),
                "quality_risk": dict(sorted(risk_counts.items())),
            },
            "companies": companies,
            "regions": regions,
            "plants": plants,
            "mapping": {
                "method_counts": self._counts(solar["weather_mapping_method"]),
                "status_counts": self._counts(solar["model_ready_status"]),
                "location_basis_counts": dict(
                    sorted(Counter(plant["location_basis"] for plant in plants).items())
                ),
                "validation": validation,
                "gold_consistency": gold_consistency,
            },
            "data_inventory": self._inventory(),
            "feature_contract": {
                "features": manifest.get("features", []),
                "energy_source_filter": manifest.get("energy_source_filter"),
                "quality_policy": manifest.get("quality_policy", {}),
                "weather_availability": manifest.get("weather_availability", {}),
            },
        }

    def _publish_static_assets(self) -> None:
        """Copy the canonical dashboard shell when publishing elsewhere."""

        source_root = (self.project_root / "dashboard").resolve()
        if source_root == self.output_dir:
            return
        relative_paths = (
            Path("solar_dashboard.html"),
            Path("plant_region_report_perm.html"),
            Path("assets/dashboard.css"),
            Path("assets/dashboard.js"),
        )
        for relative_path in relative_paths:
            source = source_root / relative_path
            if not source.is_file():
                raise FileNotFoundError(f"Dashboard static asset is missing: {source}")
            target = self.output_dir / relative_path
            target.parent.mkdir(parents=True, exist_ok=True)
            temporary = target.with_name(target.name + ".part")
            temporary.write_bytes(source.read_bytes())
            temporary.replace(target)

    def _copy_province_boundaries(
        self, national_inventory: dict[str, Any]
    ) -> Path | None:
        """Publish the local 17-province GeoJSON beside the generated payload."""

        source_label = national_inventory.get("source", {}).get(
            "boundary_path", "map/json/geoJson.json"
        )
        source = Path(str(source_label))
        if not source.is_absolute():
            source = self.project_root / source
        if not source.is_file():
            return None
        target = self.output_dir / "data/korea_provinces.geojson"
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = target.with_suffix(".geojson.part")
        boundaries = json.loads(source.read_text(encoding="utf-8"))
        temporary.write_text(
            json.dumps(boundaries, ensure_ascii=False, indent=2, allow_nan=False),
            encoding="utf-8",
        )
        temporary.replace(target)
        return target

    def _load_weather_stations(self) -> dict[int, dict[str, Any]]:
        path = self.project_root / "file/KMA_data_file/META_관측지점정보.csv"
        if not path.exists():
            return {}
        frame = read_csv_with_fallback(path)
        frame = frame.sort_values("시작일", kind="stable").drop_duplicates("지점", keep="last")
        result: dict[int, dict[str, Any]] = {}
        for _, row in frame.iterrows():
            station_id = self._integer(row.get("지점"))
            latitude = self._number(row.get("위도"))
            longitude = self._number(row.get("경도"))
            if station_id is None or not self._valid_coordinates(latitude, longitude):
                continue
            result[station_id] = {
                "name": self._text(row.get("지점명")),
                "latitude": latitude,
                "longitude": longitude,
            }
        return result

    def _plant_record(
        self,
        row: pd.Series,
        stations: dict[int, dict[str, Any]],
    ) -> dict[str, Any]:
        latitude = self._number(row.get("latitude"))
        longitude = self._number(row.get("longitude"))
        location_basis = "plant_coordinate"
        if not self._valid_coordinates(latitude, longitude):
            station_id = self._integer(row.get("weather_station_id"))
            station = stations.get(station_id or -1)
            if station:
                latitude = station["latitude"]
                longitude = station["longitude"]
                location_basis = "weather_station_proxy"
            else:
                latitude = None
                longitude = None
                location_basis = "unmapped"
        return {
            "plant_id": self._text(row.get("plant_id")),
            "company": self._text(row.get("company")),
            "company_name": COMPANY_NAMES.get(str(row.get("company")), str(row.get("company"))),
            "plant": self._text(row.get("plant")),
            "status": self._text(row.get("model_ready_status")),
            "reason": self._text(row.get("model_ready_reason")),
            "capacity_mw": self._number(row.get("capacity_mw")),
            "admin_province": self._text(row.get("admin_province")) or "행정구역 미확인",
            "admin_city": self._text(row.get("admin_city")),
            "weather_station_id": self._integer(row.get("weather_station_id")),
            "weather_station_name": self._text(row.get("weather_station_name")),
            "weather_mapping_method": self._text(row.get("weather_mapping_method")),
            "weather_mapping_confidence": self._text(row.get("weather_mapping_confidence")),
            "latitude": latitude,
            "longitude": longitude,
            "location_basis": location_basis,
            "generation_start": self._text(row.get("generation_start")),
            "generation_end": self._text(row.get("generation_end")),
            "observation_rows": self._integer(row.get("source_observation_rows")) or 0,
            "quality_rows": self._integer(row.get("rows")) or 0,
            "hourly_coverage": self._number(row.get("hourly_coverage")),
            "missing_generation_rate": self._number(row.get("missing_generation_rate")),
            "negative_generation_rate": self._number(row.get("negative_generation_rate")),
            "missing_weather_rate": self._number(row.get("missing_weather_rate")),
            "sensor_risk": self._text(row.get("sensor_risk")) or "not_evaluated",
            "recommended_action": self._text(row.get("recommended_action")),
        }

    def _group_summary(
        self,
        frame: pd.DataFrame,
        column: str,
        *,
        company_labels: bool = False,
    ) -> list[dict[str, Any]]:
        working = frame.copy()
        working[column] = working[column].fillna("행정구역 미확인").astype(str)
        records: list[dict[str, Any]] = []
        for name, group in working.groupby(column, sort=False):
            capacity = pd.to_numeric(group["capacity_mw"], errors="coerce")
            records.append(
                {
                    "key": name,
                    "name": COMPANY_NAMES.get(name, name) if company_labels else name,
                    "assets": int(len(group)),
                    "eligible": int(group["model_ready_status"].eq("eligible").sum()),
                    "quarantined": int(group["model_ready_status"].eq("quarantined").sum()),
                    "known_capacity_mw": round(float(capacity.sum(min_count=1)), 4)
                    if capacity.notna().any()
                    else None,
                    "capacity_known_assets": int(capacity.notna().sum()),
                    "observation_rows": int(
                        pd.to_numeric(group["source_observation_rows"], errors="coerce").fillna(0).sum()
                    ),
                }
            )
        return sorted(records, key=lambda item: (-item["eligible"], item["name"]))

    def _validate_registry(self, registry: pd.DataFrame) -> list[dict[str, Any]]:
        solar = registry.loc[registry["energy_source"].eq("solar")]
        eligible = solar.loc[solar["model_ready_status"].eq("eligible")]
        latitude = pd.to_numeric(registry["latitude"], errors="coerce")
        longitude = pd.to_numeric(registry["longitude"], errors="coerce")
        coordinate_pair_mismatch = int(latitude.isna().ne(longitude.isna()).sum())
        invalid_coordinates = int(
            ((latitude.notna() & ~latitude.between(32, 39)) | (longitude.notna() & ~longitude.between(124, 132))).sum()
        )
        checks = [
            ("plant_id_unique", not registry["plant_id"].duplicated().any(), int(registry["plant_id"].duplicated().sum())),
            ("eligible_solar_has_weather_station", eligible["weather_station_id"].notna().all(), int(eligible["weather_station_id"].isna().sum())),
            ("coordinate_pairs_complete", coordinate_pair_mismatch == 0, coordinate_pair_mismatch),
            ("coordinates_inside_korea_bounds", invalid_coordinates == 0, invalid_coordinates),
            ("unresolved_assets_are_quarantined", solar.loc[solar["weather_mapping_method"].eq("unresolved"), "model_ready_status"].eq("quarantined").all(), int(solar.loc[solar["weather_mapping_method"].eq("unresolved"), "model_ready_status"].ne("quarantined").sum())),
        ]
        return [
            {"check": name, "passed": bool(passed), "violations": violations}
            for name, passed, violations in checks
        ]

    def _audit_gold_partitions(self, registry: pd.DataFrame) -> dict[str, Any]:
        """Check millions of Gold rows one partition at a time with bounded memory."""

        paths = sorted(
            (self.project_root / "file/standardized/model_ready_parts").rglob("part.csv.gz")
        )
        if not paths:
            return {
                "available": False,
                "partitions": 0,
                "rows": 0,
                "plants": 0,
                "violations": 0,
            }
        lookup = registry.set_index("plant_id")
        columns = [
            "plant_id",
            "company",
            "plant",
            "energy_source",
            "region",
            "admin_province",
            "admin_city",
            "station_id",
            "weather_station_name",
            "weather_mapping_method",
        ]
        rows = 0
        plants: set[str] = set()
        violations = 0
        for path in paths:
            part = pd.read_csv(path, usecols=columns)
            rows += len(part)
            unique = part.drop_duplicates(columns)
            for row in unique.itertuples(index=False):
                plants.add(str(row.plant_id))
                if row.plant_id not in lookup.index:
                    violations += 1
                    continue
                expected = lookup.loc[row.plant_id]
                region = (
                    expected["admin_province"]
                    if pd.notna(expected["admin_province"])
                    else expected["admin_city"]
                    if pd.notna(expected["admin_city"])
                    else "unknown"
                )
                pairs = (
                    (row.company, expected["company"]),
                    (row.plant, expected["plant"]),
                    (row.energy_source, expected["energy_source"]),
                    (row.region, region),
                    (row.admin_province, expected["admin_province"]),
                    (row.admin_city, expected["admin_city"]),
                    (row.station_id, expected["weather_station_id"]),
                    (row.weather_station_name, expected["weather_station_name"]),
                    (row.weather_mapping_method, expected["weather_mapping_method"]),
                )
                violations += sum(not self._same(left, right) for left, right in pairs)
        return {
            "available": True,
            "partitions": len(paths),
            "rows": rows,
            "plants": len(plants),
            "violations": int(violations),
            "processing": "partition-at-a-time unique-key audit",
        }

    def _inventory(self) -> dict[str, Any]:
        roots = [
            ("retained_provider_archive", self.project_root / "file/solar_data_file"),
            ("current_bronze_and_legacy_normalized", self.project_root / "file/raw"),
        ]
        encodings: Counter[str] = Counter()
        roles: Counter[str] = Counter()
        naming: Counter[str] = Counter()
        noncanonical: list[str] = []
        total_bytes = 0
        for root_label, root in roots:
            if not root.exists():
                continue
            for path in root.rglob("*.csv"):
                relative = path.relative_to(self.project_root)
                audit = inspect_csv_artifact(path)
                total_bytes += audit.bytes
                encodings[audit.encoding] += 1
                is_derived = "normalized" in {part.lower() for part in relative.parts}
                role = "derived_legacy" if is_derived else "provider_original"
                roles[role] += 1
                is_metadata = "location" in relative.parts
                if is_metadata:
                    naming["metadata_not_applicable"] += 1
                elif is_canonical_solar_download_name(path):
                    naming["canonical"] += 1
                else:
                    naming["legacy_retained"] += 1
                    if len(noncanonical) < 50:
                        noncanonical.append(str(relative).replace("\\", "/"))
        return {
            "csv_files": int(sum(encodings.values())),
            "bytes": total_bytes,
            "encoding_counts": dict(sorted(encodings.items())),
            "role_counts": dict(sorted(roles.items())),
            "filename_counts": dict(sorted(naming.items())),
            "legacy_filename_examples": noncanonical,
            "decision": (
                "원본 파일은 해시와 공급기관 인코딩을 유지한다. 새 다운로드만 표준 명명 규칙을 적용하고, "
                "표준화 파일은 UTF-8-SIG로 별도 저장한다."
            ),
        }

    @staticmethod
    def _counts(series: pd.Series) -> dict[str, int]:
        return {
            str(key): int(value)
            for key, value in series.fillna("missing").value_counts().sort_index().items()
        }

    @staticmethod
    def _text(value: object) -> str | None:
        if value is None or pd.isna(value):
            return None
        text = str(value).strip()
        return text or None

    @staticmethod
    def _number(value: object) -> float | None:
        try:
            number = float(value)
        except (TypeError, ValueError):
            return None
        return number if pd.notna(number) else None

    @staticmethod
    def _integer(value: object) -> int | None:
        number = DashboardBuilder._number(value)
        return int(number) if number is not None else None

    @staticmethod
    def _valid_coordinates(latitude: float | None, longitude: float | None) -> bool:
        return (
            latitude is not None
            and longitude is not None
            and 32 <= latitude <= 39
            and 124 <= longitude <= 132
        )

    @staticmethod
    def _same(left: object, right: object) -> bool:
        if pd.isna(left) and pd.isna(right):
            return True
        if isinstance(left, (int, float)) and isinstance(right, (int, float)):
            return float(left) == float(right)
        return str(left) == str(right)
