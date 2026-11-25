"""Dataset commands: CLI input translation and command dispatch."""
from __future__ import annotations

import argparse
from pathlib import Path


def handle_audit_gold_command(args: argparse.Namespace) -> None:
    """Inspect existing artifacts; neither prepare-data nor a trainer is called."""
    import json
    import csv
    import sqlite3
    from solar_forecast.config_loader import PROJECT_ROOT, load_model_config
    from solar_forecast.datasets.gold_audit import audit_gold_dataset
    from solar_forecast.models.shared.checkpoint_audit import audit_training_compatibility

    try:
        report = audit_gold_dataset(Path(args.manifest), project_root=PROJECT_ROOT,
                                    batch_rows=args.batch_rows)
        report["gold_consistency_status"] = report["status"]
        report["data_admission_checked"] = False
        configs = args.config or ["config/models/cnn_bilstm.json", "config/models/xgboost.json"]
        report["training_compatibility"] = [
            audit_training_compatibility(load_model_config(Path(path)), report)
            for path in configs
        ]
        if any(item["schema_status"] == "FAIL" for item in report["training_compatibility"]):
            report["status"] = "FAIL"
        if args.json:
            print(json.dumps(report, ensure_ascii=False, indent=2))
        else:
            print(f"Gold/manifest consistency: {report['gold_consistency_status']}")
            print(f"Rows: {report['dataset']['rows']}; plants: {report['plants']}; partitions: {len(report['partitions'])}")
            print(f"Key/field differences: {report['comparison']}")
            for item in report["training_compatibility"]:
                print(f"{item['model']}: schema={item['schema_status']}; checkpoint={item['checkpoint_status']}; "
                      f"legacy_contracts={item['unsupported_checkpoint_contract_count']}")
            if report["issues"]:
                print(json.dumps(report["issues"], ensure_ascii=False))
            print(f"Overall audit: {report['status']}")
            print("No training or resume executed. Matching metadata still requires the trainer's stage signature.")
    except (ValueError, OSError, KeyError, csv.Error, sqlite3.Error) as exc:
        raise SystemExit(f"audit-gold: {exc}") from exc
    if report["status"] != "PASS":
        raise SystemExit(2)


def handle_prepare_data_command(args: argparse.Namespace) -> None:
    from solar_forecast.collectors import KrcYeongamCandidateIntakeService
    from solar_forecast.preprocessing.preparation_service import DataPreparationService

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


def handle_training_data_eligibility_command(args: argparse.Namespace) -> None:
    """Audit Gold plant-period continuity/split sufficiency without training."""
    from solar_forecast.preprocessing.training_eligibility import run_training_eligibility_audit

    report = run_training_eligibility_audit(
        Path(args.config),
        data_path=Path(args.data) if args.data else None,
        output_path=Path(args.output) if args.output else None,
    )
    population = report["population"]
    print(
        "Training-data eligibility: "
        f"{population['candidate_plants']} structurally eligible, "
        f"{population['structurally_rejected_plants']} structural rejects, "
        f"{population['quality_eligible_target_rows']} quality-eligible target rows"
    )
    print(f"Structural rules applied: {str(report['structural_rules_applied']).lower()}")
    print("Fixed minimum-row threshold used: false")
    print("Training admission freeze is the next preprocessing gate.")
    print(f"Next decision: {report['next_decision']}")
    if args.output:
        print(f"Eligibility manifest: {args.output}")
