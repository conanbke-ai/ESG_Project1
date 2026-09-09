"""수집·데이터 준비·학습·대시보드 연결의 단계별 실행 결과 검증."""
from __future__ import annotations

from solar_forecast.infrastructure.project_paths import COLLECTOR_SILVER_ROOT, GENERATION_ARCHIVE_ROOT, LEGACY_MERGED_SOURCE, SILVER_ROOT, VERIFICATION_ROOT, WEATHER_ARCHIVE_ROOT

from dataclasses import dataclass, field
from datetime import date, datetime, timezone
import json
from pathlib import Path
from time import perf_counter
from typing import Any, Callable

from solar_forecast.infrastructure.artifact_store import write_json_atomic
from solar_forecast.collectors import CollectionConfig, CollectionService
from solar_forecast.jobs.contracts import (
    E2E_VERIFICATION_MANIFEST_CONTRACT,
    JOB_CONTRACT_SCHEMA_VERSION,
)
from solar_forecast.jobs.training_job import TrainingService
from solar_forecast.datasets.preparation_service import DataPreparationService
from solar_forecast.reporting import DashboardBuilder
from solar_forecast.config_loader import ModelJobConfig, PROJECT_ROOT, load_model_config


TERMINAL_COLLECTION_FAILURES = {"failed", "configuration_required", "unsupported"}


@dataclass(frozen=True)
class VerificationConfig:
    """Runtime contract for a bounded end-to-end pipeline verification."""

    project_root: Path = PROJECT_ROOT
    report_root: Path = VERIFICATION_ROOT
    collect: bool = False
    collection_config: CollectionConfig | None = None
    collection_manifest: Path | None = None
    allow_collection_failures: bool = False
    prepare_data: bool = True
    input_root: Path = GENERATION_ARCHIVE_ROOT
    weather_root: Path = WEATHER_ARCHIVE_ROOT
    merged_source: Path = LEGACY_MERGED_SOURCE
    standardized_output_dir: Path = SILVER_ROOT
    collected_generation_dir: Path | None = COLLECTOR_SILVER_ROOT
    train_models: tuple[str, ...] = ("xgboost", "cnn_bilstm")
    smoke: bool = True
    model_config_paths: dict[str, Path] = field(default_factory=dict)
    no_optuna: bool = False
    max_trials: int | None = None
    optimizer_timeout_seconds: int | None = None
    build_dashboard: bool = True
    dashboard_output_dir: Path = Path("dashboard")


@dataclass(frozen=True)
class PipelineVerificationResult:
    status: str
    report_path: Path
    run_dir: Path
    steps: tuple[dict[str, Any], ...]


class PipelineVerificationService:
    """Run the public data-to-model wiring in a reproducible, auditable way.

    The default mode intentionally uses model smoke training.  It verifies file
    loading, column contracts, temporal splitting, checkpoint wiring, prediction
    artifact creation and dashboard publishing without polluting the formal
    model-comparison dashboard with short diagnostic runs.
    """

    def __init__(self, config: VerificationConfig):
        self.config = config
        self.project_root = Path(config.project_root).resolve()

    def run(self) -> PipelineVerificationResult:
        run_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        run_dir = self._resolve(self.config.report_root) / run_id
        report_path = run_dir / "verification_report.json"
        started = datetime.now(timezone.utc)
        steps: list[dict[str, Any]] = []
        status = "passed"
        error: BaseException | None = None

        try:
            self._run_step(steps, "collection_manifest", self._inspect_collection_manifest)
            if self.config.collect:
                self._run_step(steps, "official_collection", self._run_collection)
                self._run_step(
                    steps,
                    "collection_manifest_after_collect",
                    self._inspect_collection_manifest,
                )
            if self.config.prepare_data:
                self._run_step(steps, "prepare_data", self._run_prepare_data)
            for model in self.config.train_models:
                self._run_step(
                    steps,
                    f"train_{model}",
                    lambda model=model: self._run_train(model),
                )
            if self.config.build_dashboard:
                self._run_step(steps, "build_dashboard", self._run_dashboard)
        except BaseException as exc:
            status = "failed"
            error = exc
        else:
            if self._warning_count(steps):
                status = "passed_with_warnings"
        finally:
            payload = {
                "status": status,
                "contract": E2E_VERIFICATION_MANIFEST_CONTRACT,
                "schema_version": JOB_CONTRACT_SCHEMA_VERSION,
                "started_at_utc": started.isoformat(),
                "finished_at_utc": datetime.now(timezone.utc).isoformat(),
                "project_root": str(self.project_root),
                "execution": {
                    "training_mode": "smoke" if self.config.smoke else "full",
                    "formal_dashboard_comparison_includes_this_run": not self.config.smoke,
                    "models": list(self.config.train_models),
                },
                "summary": {
                    "steps": len(steps),
                    "failed_steps": sum(step["status"] == "failed" for step in steps),
                    "warnings": self._warning_count(steps),
                },
                "steps": steps,
            }
            write_json_atomic(report_path, payload)

        result = PipelineVerificationResult(status, report_path, run_dir, tuple(steps))
        if error is not None:
            raise RuntimeError(
                f"E2E verification failed at {report_path}: {error}"
            ) from error
        return result

    def _run_step(
        self,
        steps: list[dict[str, Any]],
        name: str,
        action: Callable[[], dict[str, Any]],
    ) -> None:
        started = perf_counter()
        error: BaseException | None = None
        details: dict[str, Any]
        status = "passed"
        try:
            details = action()
        except BaseException as exc:
            status = "failed"
            details = {"error": str(exc), "error_type": exc.__class__.__name__}
            error = exc
        steps.append(
            {
                "name": name,
                "status": status,
                "duration_seconds": round(perf_counter() - started, 3),
                "details": details,
            }
        )
        if error is not None:
            raise error

    def _inspect_collection_manifest(self) -> dict[str, Any]:
        manifest_path = (
            self._resolve(self.config.collection_manifest)
            if self.config.collection_manifest
            else self._latest_collection_manifest()
        )
        if manifest_path is None or not manifest_path.exists():
            return {
                "status": "not_found",
                "warnings": [
                    "No prior collection manifest was found. Run with --collect "
                    "to include official download verification."
                ],
            }
        payload = self._read_json(manifest_path)
        results = payload.get("results", [])
        failures = [
            row
            for row in results
            if str(row.get("status")) in TERMINAL_COLLECTION_FAILURES
        ]
        return {
            "manifest_path": str(manifest_path),
            "start_date": payload.get("start_date"),
            "end_date": payload.get("end_date"),
            "sources": [
                {
                    "source": row.get("source"),
                    "status": row.get("status"),
                    "rows": row.get("rows"),
                    "files": len(row.get("files", [])),
                    "message": row.get("message"),
                }
                for row in results
            ],
            "file_artifacts": len(payload.get("file_artifacts", [])),
            "warnings": [
                "Latest collection manifest contains failed or unsupported sources: "
                + ", ".join(str(row.get("source")) for row in failures)
            ]
            if failures
            else [],
        }

    def _run_collection(self) -> dict[str, Any]:
        if self.config.collection_config is None:
            raise ValueError("collection_config is required when collect=True")
        results = CollectionService(self.config.collection_config).run()
        failures = [
            result
            for result in results
            if result.status in TERMINAL_COLLECTION_FAILURES
        ]
        if failures and not self.config.allow_collection_failures:
            failed = ", ".join(f"{item.source}:{item.status}" for item in failures)
            raise RuntimeError(f"Official collection failed: {failed}")
        return {
            "sources": [
                {
                    "source": result.source,
                    "status": result.status,
                    "rows": result.rows,
                    "files": [str(path) for path in result.files],
                    "message": result.message,
                }
                for result in results
            ],
            "warnings": [
                "Official collection had failed sources but verification continued "
                "because allow_collection_failures=true."
            ]
            if failures
            else [],
        }

    def _run_prepare_data(self) -> dict[str, Any]:
        result = DataPreparationService(
            input_root=self._resolve(self.config.input_root),
            weather_root=self._resolve(self.config.weather_root),
            merged_source=self._resolve(self.config.merged_source),
            output_dir=self._resolve(self.config.standardized_output_dir),
            collected_generation_dir=(
                self._resolve(self.config.collected_generation_dir)
                if self.config.collected_generation_dir is not None
                else None
            ),
        ).run()
        manifest_path = result.model_dataset.path.with_name("model_ready_manifest.json")
        manifest = self._read_json(manifest_path) if manifest_path.exists() else {}
        warnings: list[str] = []
        downloads_root = (
            self._resolve(self.config.collected_generation_dir)
            if self.config.collected_generation_dir is not None
            else None
        )
        if downloads_root and downloads_root.exists() and not result.collector_admission:
            warnings.append(
                "Collected Silver download files are archived separately under "
                f"{downloads_root}; prepare-data currently rebuilds the Gold model "
                f"dataset from {self._resolve(self.config.input_root)} plus staged "
                "candidate intake rules."
            )
        collector_admission = None
        if result.collector_admission is not None:
            rejected_reasons: dict[str, int] = {}
            for item in result.collector_admission.files:
                if item.status != "rejected":
                    continue
                reason = str(item.reason or "unknown")
                reason_key = reason.split(":", 1)[0]
                rejected_reasons[reason_key] = rejected_reasons.get(reason_key, 0) + 1
            collector_admission = {
                "manifest_path": str(result.collector_admission.manifest_path),
                "source_dir": str(result.collector_admission.source_dir),
                "accepted_files": result.collector_admission.accepted_count,
                "rejected_files": result.collector_admission.rejected_count,
                "accepted_rows": result.collector_admission.accepted_rows,
                "rejected_reasons": rejected_reasons,
            }
            if result.collector_admission.rejected_count:
                warnings.append(
                    f"{result.collector_admission.rejected_count} collector Silver files "
                    "were kept out of Gold because they do not satisfy the plant-hour "
                    "generation contract."
                )
        return {
            "generation_manifest": str(result.generation.manifest_path),
            "generation_partitions": len(result.generation.partitions),
            "generation_rows": result.generation.rows,
            "model_dataset": str(result.model_dataset.path),
            "model_rows": result.model_dataset.rows,
            "model_plants": result.model_dataset.plants,
            "partitioned_dataset": str(result.model_dataset.partitions_dir)
            if result.model_dataset.partitions_dir
            else None,
            "quality_report": str(result.quality.report_path),
            "quality": {
                "high_risk_plants": result.quality.high_risk_plants,
                "review_plants": result.quality.review_plants,
                "preprocessing_artifact_plants": result.quality.preprocessing_artifact_plants,
            },
            "training_eligibility": manifest.get("training_eligibility", {}),
            "collector_admission": collector_admission,
            "source": {
                "input_root": str(self._resolve(self.config.input_root)),
                "weather_root": str(self._resolve(self.config.weather_root)),
                "merged_source": str(self._resolve(self.config.merged_source)),
                "collected_download_root": str(downloads_root)
                if downloads_root is not None
                else None,
            },
            "warnings": warnings,
        }

    def _run_train(self, model: str) -> dict[str, Any]:
        if model not in {"xgboost", "cnn_bilstm"}:
            raise ValueError(f"Unsupported verification model: {model}")
        config = self._model_config(model)
        run_dir = TrainingService().run(config, smoke=self.config.smoke)
        manifest_path = run_dir / "manifest.json"
        manifest = self._read_json(manifest_path)
        if manifest.get("status") != "completed":
            raise RuntimeError(f"{model} run did not complete: {manifest_path}")
        details = manifest.get("details", {})
        warnings = []
        if self.config.smoke:
            warnings.append(
                "Smoke training validates wiring only and is intentionally excluded "
                "from the formal dashboard model-comparison table."
            )
        return {
            "run_dir": str(run_dir),
            "manifest_path": str(manifest_path),
            "model": model,
            "status": manifest.get("status"),
            "metrics": details.get("metrics", {}),
            "temporal_split": details.get("temporal_split", {}),
            "memory_aware_loading": details.get("memory_aware_loading", {}),
            "checkpoint": details.get("checkpoint", {}),
            "prediction_artifacts": self._prediction_artifacts(details),
            "warnings": warnings,
        }

    def _run_dashboard(self) -> dict[str, Any]:
        result = DashboardBuilder(
            self.project_root,
            self._resolve(self.config.dashboard_output_dir),
        ).build()
        warnings = []
        if result.model_analysis_status == "empty":
            warnings.append(
                "Dashboard model analysis is empty because no compatible full "
                "XGBoost/CNN-BiLSTM evaluation pair is available. Smoke runs are "
                "kept out of public model comparison by design."
            )
        return {
            "data_path": str(result.data_path),
            "boundary_path": str(result.boundary_path),
            "solar_dashboard": str(result.solar_dashboard),
            "forecast_dashboard": str(result.forecast_dashboard),
            "model_analysis_dashboard": str(result.analytics_dashboard),
            "national_generator_records": result.national_generator_records,
            "national_capacity_mw": result.national_capacity_mw,
            "model_analysis_status": result.model_analysis_status,
            "data_quality_signals": result.data_quality_signals,
            "warnings": warnings,
        }

    def _model_config(self, model: str) -> ModelJobConfig:
        config_path = self.config.model_config_paths.get(
            model, Path(f"config/models/{model}.json")
        )
        config = load_model_config(self._resolve(config_path))
        if config.model != model:
            raise ValueError(f"Config model '{config.model}' does not match '{model}'")
        values = dict(config.values)
        optimizer = dict(values.get("optimizer", {}))
        if self.config.no_optuna:
            optimizer["enabled"] = False
        if self.config.max_trials is not None:
            optimizer["max_trials"] = self.config.max_trials
        if self.config.optimizer_timeout_seconds is not None:
            optimizer["timeout_seconds"] = self.config.optimizer_timeout_seconds
        values["optimizer"] = optimizer
        return ModelJobConfig(config.model, config.profile, values, config.source)

    def _latest_collection_manifest(self) -> Path | None:
        runs_root = self.project_root / "file" / "raw" / "runs"
        if not runs_root.exists():
            return None
        manifests = sorted(runs_root.glob("*/collection_manifest.json"))
        return manifests[-1] if manifests else None

    def _resolve(self, path: str | Path | None) -> Path:
        if path is None:
            raise ValueError("path cannot be None")
        candidate = Path(path)
        return candidate if candidate.is_absolute() else self.project_root / candidate

    @staticmethod
    def _read_json(path: Path) -> dict[str, Any]:
        return json.loads(Path(path).read_text(encoding="utf-8"))

    @staticmethod
    def _warning_count(steps: list[dict[str, Any]]) -> int:
        return sum(
            len(step.get("details", {}).get("warnings", []))
            for step in steps
            if isinstance(step.get("details"), dict)
        )

    @staticmethod
    def _prediction_artifacts(details: dict[str, Any]) -> dict[str, str]:
        return {
            key: str(value)
            for key, value in details.items()
            if key.endswith("_predictions_path") and value
        }


def build_collection_config(
    *,
    start_date: str,
    end_date: str | None,
    sources: str,
    output_dir: str | Path,
    standardized_output_dir: str | Path,
    overwrite: bool,
    download_date: str | None,
    komipo_station_codes: tuple[str, ...],
    api_max_calls: int,
    station_ids: tuple[str, ...] = (),
    kma_mode: str = "auto",
    weather_root: str | Path = "file/KMA_data_file",
) -> CollectionConfig:
    """Create a collection config for the verification CLI boundary."""

    return CollectionConfig(
        start_date=date.fromisoformat(start_date),
        end_date=date.fromisoformat(end_date) if end_date else date.today(),
        sources=tuple(item.strip() for item in sources.split(",") if item.strip()),
        output_dir=Path(output_dir),
        standardized_output_dir=Path(standardized_output_dir),
        overwrite=overwrite,
        download_date=date.fromisoformat(download_date) if download_date else date.today(),
        komipo_station_codes=komipo_station_codes,
        api_max_calls=api_max_calls,
        station_ids=station_ids,
        kma_mode=kma_mode,
        existing_weather_dir=Path(weather_root),
    )


__all__ = [
    "PipelineVerificationResult",
    "PipelineVerificationService",
    "VerificationConfig",
    "build_collection_config",
]
