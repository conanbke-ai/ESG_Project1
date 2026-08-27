from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from solar_forecast.collectors.archive import (
    HistoricalGenerationStandardizationService,
    StandardizationRun,
)
from solar_forecast.collectors.candidates import (
    CandidateIntakeResult,
    KrcYeongamCandidateIntakeService,
)
from solar_forecast.collectors.metadata import PlantMetadataCatalog
from solar_forecast.features.service import ModelDatasetResult, NationwideModelDatasetBuilder
from solar_forecast.pipeline.dataset import DatasetLoadPolicy, DatasetRepository
from solar_forecast.quality import QualityAuditResult, QualityAuditService
from solar_forecast.quality.policy import WEATHER_RANGES


@dataclass(frozen=True)
class DataPreparationResult:
    generation: StandardizationRun
    model_dataset: ModelDatasetResult
    quality: QualityAuditResult
    legacy_quality: QualityAuditResult
    candidate_intake: CandidateIntakeResult | None = None


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
        )
        candidate_intake: CandidateIntakeResult | None = None
        krc_source = self.input_root.parent / "raw" / "한국농어촌공사" / "영암"
        if krc_source.exists() and any(krc_source.glob("*.csv")):
            candidate_intake = KrcYeongamCandidateIntakeService(
                krc_source,
                self.output_dir / "candidates" / "krc_yeongam",
            ).run()
            if candidate_intake.status == "generation_ready_for_registry":
                candidate_paths = tuple(
                    sorted(
                        (self.output_dir / "candidates" / "krc_yeongam" / "generation").glob(
                            "*.csv.gz"
                        )
                    )
                )
                generation_paths = (*generation_paths, *candidate_paths)
        builder = NationwideModelDatasetBuilder(self.weather_root, metadata)
        legacy_generation = builder.read_legacy_generation(self.merged_source)
        legacy_quality = QualityAuditService().run(
            legacy_generation,
            self.output_dir,
            reference_paths=generation_paths,
            artifact_name="legacy_pipeline_quality",
        )
        model_dataset = builder.build(
            self.merged_source,
            self.output_dir / "model_ready.csv.gz",
            generation_paths=generation_paths,
        )
        quality_context = [
            "timestamp",
            "company",
            "plant_id",
            "plant",
            "region",
            "energy_source",
            "generation_mwh",
            "capacity_mw",
            *WEATHER_RANGES,
        ]
        _, model_frame, _ = DatasetRepository(model_dataset.path.parent).load_training_frame(
            model_dataset.partitions_dir or model_dataset.path,
            columns=quality_context,
            numeric_columns=["generation_mwh", "capacity_mw", *WEATHER_RANGES],
            policy=DatasetLoadPolicy(chunk_rows=100_000, memory_limit_mb=1536),
        )
        quality = QualityAuditService().run(
            model_frame,
            self.output_dir,
            reference_paths=generation_paths,
        )
        return DataPreparationResult(
            generation,
            model_dataset,
            quality,
            legacy_quality,
            candidate_intake,
        )
