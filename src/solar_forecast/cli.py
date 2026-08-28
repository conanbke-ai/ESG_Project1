from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
from typing import Sequence

from solar_forecast.collectors import (
    CollectionConfig,
    CollectionService,
    KrcYeongamCandidateIntakeService,
)
from solar_forecast.ensemble.service import HybridExperiment
from solar_forecast.evaluation import FeatureAblationService
from solar_forecast.jobs.training import TrainingService
from solar_forecast.preparation import DataPreparationService
from solar_forecast.settings import ModelJobConfig, PROJECT_ROOT, load_model_config


def _csv_list(value: str | None) -> list[str] | None:
    return [item.strip() for item in value.split(",") if item.strip()] if value else None


def _sequence(args: argparse.Namespace):
    from solar_forecast.models.cnn.config import SequenceConfig

    return SequenceConfig(
        sequence_length=args.sequence_length,
        test_size=args.test_size,
        val_size=args.val_size,
        calibration_size=args.calibration_size,
        purge_gap_hours=args.purge_gap_hours,
        batch_size=args.batch_size,
        shuffle=not args.no_shuffle,
        append_missing_indicators=not args.no_missing_indicators,
        num_workers=args.num_workers,
    )


def _add_sequence_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--sequence-length", type=int, default=24)
    parser.add_argument("--test-size", type=float, default=0.15)
    parser.add_argument("--val-size", type=float, default=0.15)
    parser.add_argument("--calibration-size", type=float, default=0.10)
    parser.add_argument("--purge-gap-hours", type=int, default=0)
    parser.add_argument("--batch-size", type=int, default=64)
    parser.add_argument(
        "--no-shuffle",
        action="store_true",
        help="Disable Train batch shuffling; temporal split membership is never shuffled",
    )
    parser.add_argument("--no-missing-indicators", action="store_true")
    parser.add_argument("--num-workers", type=int, default=0)


def _run_pipeline(args: argparse.Namespace) -> None:
    from solar_forecast.pipeline import ForecastPipeline, PipelineConfig

    config = PipelineConfig(
        target_column=args.target,
        data_path=Path(args.data) if args.data else None,
        input_dir=Path(args.input_dir),
        feature_columns=_csv_list(args.features),
        output_dir=Path(args.output_dir),
        sequence=_sequence(args),
        epochs=args.epochs,
        n_trials=args.n_trials,
        use_optuna=not args.no_optuna,
        optimizer_timeout_seconds=args.optimizer_timeout_seconds,
        use_reinforcement=args.reinforcement,
        contamination=args.contamination,
        artifact_level=args.artifact_level,
    )
    result = ForecastPipeline(config).run()
    print(f"Pipeline completed: {result.run_dir}")
    print(f"HTML report: {result.report_path}")


def _run_hybrid(args: argparse.Namespace) -> None:
    paths = HybridExperiment(Path(args.output_dir), args.artifact_level).run(
        Path(args.validation), Path(args.test)
    )
    print(f"Hybrid completed: {args.output_dir}")
    print(f"National metrics: {paths['national_metrics']}")


def _run_collect(args: argparse.Namespace) -> None:
    config = CollectionConfig(
        start_date=date.fromisoformat(args.start_date),
        end_date=date.fromisoformat(args.end_date) if args.end_date else date.today(),
        sources=_csv_list(args.sources) or [],
        output_dir=Path(args.output_dir),
        standardized_output_dir=Path(args.standardized_output_dir),
        overwrite=args.overwrite,
        komipo_station_codes=_csv_list(args.komipo_station_codes) or [],
        api_max_calls=args.api_max_calls,
        download_date=date.fromisoformat(args.download_date) if args.download_date else date.today(),
    )
    results = CollectionService(config).run()
    for result in results:
        print(f"{result.source}: {result.status} ({len(result.files)} files, {result.rows} rows) {result.message}")
    if any(result.status in {"failed", "configuration_required", "unsupported"} for result in results):
        raise SystemExit(1)


def _run_train(args: argparse.Namespace) -> None:
    config_path = Path(args.config) if args.config else Path(f"config/models/{args.model}.json")
    config = load_model_config(config_path)
    if config.model != args.model:
        raise ValueError(f"Config model '{config.model}' does not match '{args.model}'")
    values = dict(config.values)
    optimizer = dict(values.get("optimizer", {}))
    if args.no_optuna:
        optimizer["enabled"] = False
    if args.max_trials is not None:
        optimizer["max_trials"] = args.max_trials
    if args.optimizer_timeout_seconds is not None:
        optimizer["timeout_seconds"] = args.optimizer_timeout_seconds
    values["optimizer"] = optimizer
    config = ModelJobConfig(config.model, config.profile, values, config.source)
    run_dir = TrainingService().run(config, smoke=args.smoke)
    print(f"Completed {args.model}: {run_dir}")


def _run_prepare_data(args: argparse.Namespace) -> None:
    result = DataPreparationService(
        input_root=Path(args.input_root),
        weather_root=Path(args.weather_root),
        merged_source=Path(args.merged_source),
        output_dir=Path(args.output_dir),
    ).run()
    print(
        f"Generation standardized: {len(result.generation.partitions)} files, "
        f"{result.generation.rows} hourly rows"
    )
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


def _run_audit_candidate_data(args: argparse.Namespace) -> None:
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


def _run_status(_: argparse.Namespace) -> None:
    lock = PROJECT_ROOT / "artifacts" / ".training.lock"
    print(f"Training active: {lock.read_text(encoding='utf-8')}" if lock.exists() else "No training job is active")


def _run_evaluate_features(args: argparse.Namespace) -> None:
    result = FeatureAblationService(n_estimators=args.n_estimators).run(
        Path(args.data),
        Path(args.output_dir),
        n_splits=args.folds,
        validation_window_hours=args.validation_window_hours,
        calibration_fraction=args.calibration_fraction,
        test_fraction=args.test_fraction,
        gap_hours=args.gap_hours,
    )
    print(f"Feature ablation: {result.result_path}")
    print(f"Selected contract: {result.selected_contract} ({len(result.selected_features)} features)")


def _run_build_dashboard(args: argparse.Namespace) -> None:
    from solar_forecast.reporting import DashboardBuilder

    result = DashboardBuilder(PROJECT_ROOT, Path(args.output_dir)).build()
    print(
        f"Dashboard data refreshed: {result.solar_assets} solar assets, "
        f"{result.eligible_solar_assets} eligible"
    )
    print(f"Solar dashboard: {result.solar_dashboard}")
    print(f"Plant/region report: {result.mapping_report}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Solar forecast and anomaly monitoring")
    commands = parser.add_subparsers(dest="command", required=True)

    pipeline = commands.add_parser("pipeline", help="Run preprocessing, CNN training, analysis, and reporting")
    pipeline.add_argument("target")
    pipeline.add_argument("--data")
    pipeline.add_argument("--input-dir", default="file/merge_data")
    pipeline.add_argument("--features")
    pipeline.add_argument("--output-dir", default="output/pipeline")
    pipeline.add_argument("--epochs", type=int, default=50)
    pipeline.add_argument("--n-trials", type=int, default=10)
    pipeline.add_argument("--optimizer-timeout-seconds", type=int)
    pipeline.add_argument("--no-optuna", action="store_true")
    pipeline.add_argument("--reinforcement", action="store_true")
    pipeline.add_argument("--contamination", type=float, default=0.05)
    pipeline.add_argument("--artifact-level", choices=["minimal", "standard", "debug"], default="minimal")
    _add_sequence_options(pipeline)
    pipeline.set_defaults(func=_run_pipeline)

    hybrid = commands.add_parser("hybrid", help="Run the explainable dynamic hybrid")
    hybrid.add_argument("validation")
    hybrid.add_argument("test")
    hybrid.add_argument("--output-dir", default="output/experiments/hybrid")
    hybrid.add_argument("--artifact-level", choices=["minimal", "standard", "debug"], default="minimal")
    hybrid.set_defaults(func=_run_hybrid)

    collect = commands.add_parser("collect", help="Collect official generation and KMA data")
    collect.add_argument("--start-date", required=True)
    collect.add_argument("--end-date")
    collect.add_argument("--sources", default="koen,kospo,ewp,iwest,kma")
    collect.add_argument("--output-dir", default="file/raw")
    collect.add_argument("--standardized-output-dir", default="file/standardized/downloads")
    collect.add_argument(
        "--download-date",
        help="Override the YYYY-MM-DD collection date embedded in canonical filenames",
    )
    collect.add_argument("--overwrite", action="store_true")
    collect.add_argument("--komipo-station-codes")
    collect.add_argument("--api-max-calls", type=int, default=900)
    collect.set_defaults(func=_run_collect)

    prepare = commands.add_parser(
        "prepare-data",
        help="Standardize retained public-provider files and build leakage-safe model features",
    )
    prepare.add_argument("--input-root", default="file/solar_data_file")
    prepare.add_argument("--weather-root", default="file/KMA_data_file")
    prepare.add_argument("--merged-source", default="file/merge_data/val.csv")
    prepare.add_argument("--output-dir", default="file/standardized")
    prepare.set_defaults(func=_run_prepare_data)

    audit_candidate = commands.add_parser(
        "audit-candidate-data",
        help="Normalize and quality-gate staged additional generation data",
    )
    audit_candidate.add_argument(
        "--source-dir",
        default="file/raw/한국농어촌공사/영암",
    )
    audit_candidate.add_argument(
        "--output-dir",
        default="file/standardized/candidates/krc_yeongam",
    )
    audit_candidate.set_defaults(func=_run_audit_candidate_data)

    train = commands.add_parser("train", help="Run one independently locked model job")
    train.add_argument("model", choices=["xgboost", "cnn_bilstm"])
    train.add_argument("--config")
    train.add_argument("--smoke", action="store_true")
    train.add_argument(
        "--no-optuna",
        action="store_true",
        help="Skip hyperparameter search and use the fixed model config",
    )
    train.add_argument(
        "--max-trials",
        type=int,
        help="Override the maximum total trials in the resumable study",
    )
    train.add_argument(
        "--optimizer-timeout-seconds",
        type=int,
        help="Override the wall-time budget for this optimization call",
    )
    train.set_defaults(func=_run_train)

    status = commands.add_parser("status", help="Show the active training job")
    status.set_defaults(func=_run_status)

    evaluate_features = commands.add_parser(
        "evaluate-features",
        help="Compare feature contracts with purged rolling-origin validation",
    )
    evaluate_features.add_argument("--data", default="file/standardized/model_ready.csv.gz")
    evaluate_features.add_argument("--output-dir", default="output/evaluation/features")
    evaluate_features.add_argument("--folds", type=int, default=3)
    evaluate_features.add_argument("--validation-window-hours", type=int, default=2160)
    evaluate_features.add_argument("--calibration-fraction", type=float, default=0.10)
    evaluate_features.add_argument("--test-fraction", type=float, default=0.15)
    evaluate_features.add_argument("--gap-hours", type=int, default=168)
    evaluate_features.add_argument("--n-estimators", type=int, default=300)
    evaluate_features.set_defaults(func=_run_evaluate_features)

    dashboard = commands.add_parser(
        "build-dashboard",
        help="Refresh dashboard data from current registry, quality, and model manifests",
    )
    dashboard.add_argument("--output-dir", default="dashboard")
    dashboard.set_defaults(func=_run_build_dashboard)
    return parser


def main(argv: Sequence[str] | None = None) -> None:
    args = build_parser().parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
