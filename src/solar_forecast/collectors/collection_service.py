"""공급자별 실행 격리와 전체 collection manifest 기록."""
from __future__ import annotations

from datetime import datetime
import logging
from pathlib import Path

from solar_forecast.infrastructure.artifact_store import write_json_atomic
from solar_forecast.jobs.contracts import COLLECTION_MANIFEST_CONTRACT, JOB_CONTRACT_SCHEMA_VERSION
from solar_forecast.infrastructure.error_reporting import write_error_report
from solar_forecast.infrastructure.environment import load_local_env

from solar_forecast.collectors.contracts import CollectionResult
from solar_forecast.collectors.collection_config import CollectionConfig
from solar_forecast.infrastructure.csv_storage import inspect_csv_artifact
from solar_forecast.collectors.collector_factory import build_collector


logger = logging.getLogger(__name__)


class CollectionService:
    """Application service that isolates collectors and owns run artifacts."""

    def __init__(self, config: CollectionConfig):
        self.config = config
        self._factories = {
            source: lambda source=source: build_collector(source, config)
            for source in ("koen", "kospo", "ewp", "iwest", "kma", "komipo")
        }

    def run(self) -> list[CollectionResult]:
        load_local_env()
        run_dir = self.config.output_dir / "runs" / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        run_dir.mkdir(parents=True, exist_ok=False)
        results = [self._collect(source, run_dir) for source in self.config.sources]
        manifest = {
            "contract": COLLECTION_MANIFEST_CONTRACT,
            "schema_version": JOB_CONTRACT_SCHEMA_VERSION,
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
                        else "standardized_weather_silver"
                        if self.config.existing_weather_dir.resolve() in path.resolve().parents
                        else "provider_original_bronze"
                    ),
                }
                for result in results
                for path in result.files
                if path.suffix.lower() == ".csv" and path.exists()
            ],
        }
        write_json_atomic(run_dir / "collection_manifest.json", manifest)
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
