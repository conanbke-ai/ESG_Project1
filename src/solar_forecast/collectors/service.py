from __future__ import annotations

from datetime import datetime
import json
import logging
from pathlib import Path

from .base import CollectionResult
from .config import CollectionConfig, load_source_catalog
from .csv_artifacts import inspect_csv_artifact
from .generation import (
    DataGoFileCollector,
    DataGoDatasetSpec,
    EwpAttachmentSpec,
    EwpTrainingDataCollector,
    KoenHomepageCollector,
)
from .kma import KmaAsosHourlyCollector
from .openapi import KomipoRenewableCollector
from .normalization import (
    DailyWideGenerationNormalizer,
    IWEST_WIDE_SCHEMA,
    KOSPO_WIDE_SCHEMA,
)
from solar_forecast.infrastructure.error_report import write_error_report
from solar_forecast.infrastructure.local_env import load_local_env


logger = logging.getLogger(__name__)


class CollectionService:
    """Application service that isolates collectors and owns run artifacts."""

    def __init__(self, config: CollectionConfig):
        self.config = config
        catalog = load_source_catalog()
        kospo = catalog["kospo"]
        ewp = catalog["ewp"]
        iwest = catalog["iwest"]
        kospo_spec = DataGoDatasetSpec(
            "kospo",
            kospo["dataset_id"],
            kospo["detail_id"],
            kospo["organization"],
            kospo["detail_name"],
        )
        iwest_spec = DataGoDatasetSpec(
            "iwest",
            iwest["dataset_id"],
            iwest["detail_id"],
            iwest["organization"],
            iwest["detail_name"],
        )
        ewp_spec = EwpAttachmentSpec(
            detail_url=ewp["dataset_page"],
            download_url=ewp["download_url"],
            attachment_id=ewp["attachment_id"],
            order_number=ewp["order_number"],
            page_code=ewp["page_code"],
            organization=ewp["organization"],
            detail_name=ewp["detail_name"],
        )
        self._factories = {
            "koen": lambda: KoenHomepageCollector(config),
            "kospo": lambda: DataGoFileCollector(
                kospo_spec, config, DailyWideGenerationNormalizer(KOSPO_WIDE_SCHEMA)
            ),
            "ewp": lambda: EwpTrainingDataCollector(config, ewp_spec),
            "iwest": lambda: DataGoFileCollector(
                iwest_spec, config, DailyWideGenerationNormalizer(IWEST_WIDE_SCHEMA)
            ),
            "kma": lambda: KmaAsosHourlyCollector(config),
            "komipo": lambda: KomipoRenewableCollector(config),
        }

    def run(self) -> list[CollectionResult]:
        load_local_env()
        run_dir = self.config.output_dir / "runs" / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        run_dir.mkdir(parents=True, exist_ok=False)
        results = [self._collect(source, run_dir) for source in self.config.sources]
        manifest = {
            "started_at": datetime.now().isoformat(),
            "start_date": self.config.start_date.isoformat(),
            "end_date": self.config.end_date.isoformat(),
            "results": [
                {"source": r.source, "status": r.status, "rows": r.rows, "files": [str(p) for p in r.files], "message": r.message}
                for r in results
            ],
            "file_artifacts": [
                {
                    **inspect_csv_artifact(path).as_dict(),
                    "role": (
                        "standardized_silver"
                        if self.config.standardized_output_dir in path.parents
                        else "provider_original_bronze"
                    ),
                }
                for result in results
                for path in result.files
                if path.suffix.lower() == ".csv" and path.exists()
            ],
        }
        (run_dir / "collection_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        return results

    def _collect(self, source: str, run_dir: Path) -> CollectionResult:
        factory = self._factories.get(source)
        if factory is None:
            return CollectionResult(source, "unsupported", message="Unknown source")
        try:
            logger.info("Collecting %s", source)
            return factory().collect()
        except Exception as exc:
            write_error_report(run_dir / source, exc, stage=f"collection:{source}")
            return CollectionResult(source, "failed", message=str(exc))


def collect_all(config: CollectionConfig) -> list[CollectionResult]:
    """Compatibility facade; new code should instantiate CollectionService."""
    return CollectionService(config).run()
