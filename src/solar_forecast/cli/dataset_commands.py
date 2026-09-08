"""Dataset commands: CLI input translation and command dispatch."""
from __future__ import annotations

import argparse
from pathlib import Path


def handle_prepare_data_command(args: argparse.Namespace) -> None:
    from solar_forecast.collectors import KrcYeongamCandidateIntakeService
    from solar_forecast.datasets.preparation_service import DataPreparationService

    result = DataPreparationService(
        input_root=Path(args.input_root),
        weather_root=Path(args.weather_root),
        merged_source=Path(args.merged_source),
        output_dir=Path(args.output_dir),
        collected_generation_dir=(
            None
            if args.no_collected_downloads
            else Path(args.collected_generation_dir)
        ),
    ).run()
    print(
        f"Generation standardized: {len(result.generation.partitions)} files, "
        f"{result.generation.rows} hourly rows"
    )
    if result.collector_admission:
        print(
            f"Collector Silver admission: {result.collector_admission.accepted_count} accepted, "
            f"{result.collector_admission.rejected_count} rejected, "
            f"{result.collector_admission.accepted_rows} admitted rows"
        )
        print(f"Collector admission manifest: {result.collector_admission.manifest_path}")
    if result.candidate_intake:
        print(
            f"Candidate generation admitted to registry gate: "
            f"{result.candidate_intake.rows} rows, {result.candidate_intake.plants} plants "
            f"({result.candidate_intake.status})"
        )
    print(f"Generation manifest: {result.generation.manifest_path}")
    print(
        f"Model dataset: {result.model_dataset.path} "
        f"({result.model_dataset.rows} rows, {result.model_dataset.plants} plants)"
    )
    print(
        f"Quality report: {result.quality.report_path} "
        f"({result.quality.high_risk_plants} high, {result.quality.review_plants} review, "
        f"{result.quality.preprocessing_artifact_plants} preprocessing-artifact)"
    )
    print(
        f"Legacy pipeline audit: {result.legacy_quality.report_path} "
        f"({result.legacy_quality.preprocessing_artifact_plants} preprocessing-artifact)"
    )


def handle_audit_candidate_data_command(args: argparse.Namespace) -> None:
    from solar_forecast.collectors import KrcYeongamCandidateIntakeService
    from solar_forecast.datasets.preparation_service import DataPreparationService

    result = KrcYeongamCandidateIntakeService(
        source_dir=Path(args.source_dir),
        output_dir=Path(args.output_dir),
    ).run()
    accepted = sum(item.status == "accepted_for_generation_audit" for item in result.source_files)
    quarantined = len(result.source_files) - accepted
    print(
        f"Candidate intake: {result.rows} hourly rows, {result.plants} plants, "
        f"{accepted} accepted files, {quarantined} quarantined files"
    )
    print(f"Admission status: {result.status}")
    print(f"Candidate manifest: {result.manifest_path}")
