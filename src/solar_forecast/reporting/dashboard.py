from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
from typing import Any

from solar_forecast.artifacts.manifest import replace_file_atomic
from solar_forecast.reporting.model_analytics import ModelAnalyticsService
from solar_forecast.reporting.national_inventory import build_national_inventory
from solar_forecast.reporting.province_boundaries import validate_province_boundaries


@dataclass(frozen=True)
class DashboardBuildResult:
    data_path: Path
    boundary_path: Path
    solar_dashboard: Path
    forecast_dashboard: Path
    analytics_dashboard: Path
    national_generator_records: int
    national_capacity_mw: float
    model_analysis_status: str
    data_quality_signals: int

    @property
    def mapping_report(self) -> Path:
        """Compatibility alias for callers using the retired report name."""

        return self.analytics_dashboard


class DashboardBuilder:
    """Publish nationwide inventory and model analytics as a clean public view."""

    def __init__(self, project_root: Path, output_dir: Path | None = None):
        self.project_root = Path(project_root).resolve()
        self.output_dir = (output_dir or self.project_root / "dashboard").resolve()

    def build(self) -> DashboardBuildResult:
        raw_inventory = build_national_inventory(self.project_root)[
            "national_inventory"
        ]
        national_inventory = self._public_inventory(raw_inventory)
        model_analysis = ModelAnalyticsService(self.project_root).build()
        payload = {
            "meta": {
                "generated_at": datetime.now().astimezone().isoformat(
                    timespec="seconds"
                ),
                "scope": "전국 태양광 설비 현황과 정식 모델 평가 결과",
            },
            "national_inventory": national_inventory,
            "model_analysis": model_analysis,
        }

        self._publish_static_assets()
        boundary_path = self._copy_province_boundaries(raw_inventory)
        data_path = self._write_payload(payload)
        anomaly_signals = model_analysis["anomalies"]["data_quality_signals"]
        return DashboardBuildResult(
            data_path=data_path,
            boundary_path=boundary_path,
            solar_dashboard=self.output_dir / "solar_dashboard.html",
            forecast_dashboard=self.output_dir / "forecast.html",
            analytics_dashboard=self.output_dir / "model_analysis.html",
            national_generator_records=int(
                national_inventory["summary"]["generator_records"]
            ),
            national_capacity_mw=float(
                national_inventory["summary"]["total_capacity_mw"]
            ),
            model_analysis_status=str(model_analysis["status"]),
            data_quality_signals=len(anomaly_signals),
        )

    @staticmethod
    def _public_inventory(inventory: dict[str, Any]) -> dict[str, Any]:
        """Remove hashes, local paths, encoding and audit internals from the UI payload."""

        locations = [
            {
                "region": row["region"],
                "subregion": row["subregion"],
                "generator_records": int(row["generator_records"]),
                "capacity_mw": row["capacity_mw"],
                "source_region_conflict": bool(
                    row.get("source_region_conflict", False)
                ),
            }
            for row in inventory.get("locations", [])
        ]
        local_counts: dict[str, int] = {}
        for row in locations:
            local_counts[row["region"]] = local_counts.get(row["region"], 0) + 1
        regions = [
            {
                "region": row["region"],
                "generator_records": int(row["generator_records"]),
                "capacity_mw": row["capacity_mw"],
                "local_area_count": local_counts.get(row["region"], 0),
            }
            for row in inventory.get("regions", [])
        ]
        source = inventory.get("source", {})
        summary = inventory.get("summary", {})
        return {
            "source": {
                "provider": source.get("provider"),
                "source_system": source.get("source_system"),
                "source_url": source.get("source_url"),
                "reference_date": source.get("reference_date"),
                "scope": source.get("scope"),
            },
            "summary": {
                "generator_records": int(summary.get("generator_records", 0)),
                "total_capacity_mw": summary.get("total_capacity_mw", 0.0),
                "canonical_regions": int(summary.get("canonical_regions", 0)),
                "regions_with_records": int(summary.get("regions_with_records", 0)),
                "subregions": int(summary.get("subregions", 0)),
                "source_region_conflicts": sum(
                    row["source_region_conflict"] for row in locations
                ),
            },
            "regions": regions,
            "locations": locations,
            "location_search_aliases": dict(
                inventory.get("location_search_aliases", {})
            ),
        }

    def _write_payload(self, payload: dict[str, Any]) -> Path:
        data_path = self.output_dir / "data/dashboard_data.json"
        data_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = data_path.with_suffix(".json.part")
        temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False),
            encoding="utf-8",
        )
        replace_file_atomic(temporary, data_path)
        return data_path

    def _publish_static_assets(self) -> None:
        """Copy the canonical dashboard shell when publishing elsewhere."""

        source_root = (self.project_root / "dashboard").resolve()
        if source_root == self.output_dir:
            return
        relative_paths = (
            Path("solar_dashboard.html"),
            Path("forecast.html"),
            Path("model_analysis.html"),
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
            replace_file_atomic(temporary, target)

    def _copy_province_boundaries(
        self, national_inventory: dict[str, Any]
    ) -> Path:
        source_label = national_inventory.get("source", {}).get(
            "boundary_path", "map/json/geoJson.json"
        )
        source = Path(str(source_label))
        if not source.is_absolute():
            source = self.project_root / source
        if not source.is_file():
            raise FileNotFoundError(
                f"Required province boundary source is missing: {source}"
            )
        target = self.output_dir / "data/korea_provinces.geojson"
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = target.with_suffix(".geojson.part")
        boundaries = json.loads(source.read_text(encoding="utf-8"))
        metadata = boundaries.get("metadata", {})
        validate_province_boundaries(
            boundaries,
            require_complete=str(metadata.get("dataset_id", "")).startswith("sgis_"),
        )
        temporary.write_text(
            json.dumps(
                boundaries,
                ensure_ascii=False,
                separators=(",", ":"),
                allow_nan=False,
            )
            + "\n",
            encoding="utf-8",
        )
        replace_file_atomic(temporary, target)
        return target


__all__ = ["DashboardBuildResult", "DashboardBuilder"]
