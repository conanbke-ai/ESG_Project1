from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from solar_forecast.collectors.archive import (
    HistoricalGenerationStandardizationService,
    StandardizationRun,
)
from solar_forecast.collectors.metadata import PlantMetadataCatalog
from solar_forecast.features.service import LegacyModelDatasetBuilder, ModelDatasetResult
from solar_forecast.pipeline.dataset import DatasetRepository
from solar_forecast.quality import QualityAuditResult, QualityAuditService


@dataclass(frozen=True)
class DataPreparationResult:
    generation: StandardizationRun
    model_dataset: ModelDatasetResult
    quality: QualityAuditResult
    legacy_quality: QualityAuditResult


class DataPreparationService:
    """Application boundary for historical standardization and model-ready features."""

    def __init__(
        self,
        input_root: Path,
        weather_root: Path,
        merged_source: Path,
        output_dir: Path,
    ):
        self.input_root = Path(input_root)
        self.weather_root = Path(weather_root)
        self.merged_source = Path(merged_source)
        self.output_dir = Path(output_dir)

    def run(self) -> DataPreparationResult:
        metadata = PlantMetadataCatalog.from_directory(self.input_root / "location")
        generation = HistoricalGenerationStandardizationService(
            self.input_root,
            self.output_dir,
            metadata=metadata,
        ).run()
        generation_paths = tuple(
            Path(partition.destination)
            for partition in generation.partitions
            if partition.company == "kospo"
        )
        builder = LegacyModelDatasetBuilder(self.weather_root, metadata)
        legacy_generation = builder.read_legacy_generation(self.merged_source)
        legacy_quality = QualityAuditService().run(
            legacy_generation,
            self.output_dir,
            reference_paths=generation_paths,
            artifact_name="legacy_pipeline_quality",
        )
        model_dataset = builder.build(
            self.merged_source,
            self.output_dir / "model_ready.csv",
            generation_paths=generation_paths,
        )
        _, model_frame = DatasetRepository(model_dataset.path.parent).load(model_dataset.path)
        quality = QualityAuditService().run(
            model_frame,
            self.output_dir,
            reference_paths=generation_paths,
        )
        return DataPreparationResult(generation, model_dataset, quality, legacy_quality)
