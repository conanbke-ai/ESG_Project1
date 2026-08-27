from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
from typing import Any, Optional

from .config import PipelineConfig
from .dataset import DatasetRepository
from .preprocessing import NumericPreprocessor
from .reporting import HtmlReportWriter
from .training import CnnTrainingAdapter
from .ports import DatasetPort, PreprocessorPort, ReportPort, TrainingPort
from solar_forecast.infrastructure.error_report import write_error_report


@dataclass(frozen=True)
class PipelineResult:
    run_dir: Path
    source_path: Path
    processed_path: Optional[Path]
    checkpoint_path: Path
    report_path: Path
    metrics: dict[str, Any]


class ForecastPipeline:
    """Application service composed from replaceable infrastructure adapters."""

    def __init__(
        self,
        config: PipelineConfig,
        repository: DatasetPort | None = None,
        preprocessor: PreprocessorPort | None = None,
        trainer: TrainingPort | None = None,
        reporter: ReportPort | None = None,
    ):
        self.config = config
        self.repository = repository or DatasetRepository(config.input_dir)
        # Sequence preparation owns a train-fitted imputer. Keeping missing
        # values here prevents validation/test statistics from leaking backward.
        self.preprocessor = preprocessor or NumericPreprocessor(fill_missing=False)
        self.trainer = trainer or CnnTrainingAdapter()
        self.reporter = reporter or HtmlReportWriter()

    def run(self) -> PipelineResult:
        config = self.config
        run_dir = Path(config.output_dir) / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        run_dir.mkdir(parents=True, exist_ok=False)
        stage = "ingestion"
        try:
            source, raw = self.repository.load(config.data_path)
            stage = "preprocessing"
            prepared = self.preprocessor.transform(raw, config.target_column, config.feature_columns)
            processed_path: Optional[Path] = None
            if config.artifact_level == "debug":
                processed_path = run_dir / "processed.csv"
                prepared.frame.to_csv(processed_path, index=False, encoding="utf-8-sig")

            stage = "training_and_evaluation"
            artifacts, analysis = self.trainer.execute(prepared.frame, prepared.feature_columns, config, run_dir)
            stage = "html_reporting"
            report_path = self.reporter.write(analysis["metrics"], analysis["anomalies"], source, run_dir / "report.html")
            result = PipelineResult(run_dir, source, processed_path, Path(artifacts["checkpoint_path"]), report_path, analysis["metrics"])
            if config.artifact_level in {"standard", "debug"}:
                analysis["anomalies"].to_csv(run_dir / "anomalies.csv.gz", index=False, compression="gzip")

            manifest = {
                "status": "completed", "artifact_level": config.artifact_level,
                "source_path": str(source), "processed_path": str(processed_path) if processed_path else None,
                "checkpoint_path": str(result.checkpoint_path), "report_path": str(report_path),
                "features": prepared.feature_columns, "dropped_rows": prepared.dropped_rows,
                "metrics": result.metrics, "anomaly_count": int(analysis["anomalies"]["is_outlier"].sum()),
            }
            (run_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
            return result
        except Exception as exc:
            write_error_report(
                run_dir, exc, stage=stage,
                context={"data_path": config.data_path, "input_dir": config.input_dir, "target": config.target_column},
            )
            raise


def run_pipeline(config: PipelineConfig) -> PipelineResult:
    """Compatibility facade around ForecastPipeline."""
    return ForecastPipeline(config).run()
