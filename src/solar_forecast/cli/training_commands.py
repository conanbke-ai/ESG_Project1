"""Training commands: CLI input translation and command dispatch."""
from __future__ import annotations

import argparse
from pathlib import Path
from solar_forecast.config_loader import ModelJobConfig, load_model_config
from solar_forecast.cli.arguments import parse_csv_values, build_sequence_config


def handle_pipeline_command(args: argparse.Namespace) -> None:
    from solar_forecast.models.hybrid.experiment import HybridExperiment
    from solar_forecast.evaluation import FeatureAblationService
    from solar_forecast.jobs.training_job import TrainingService

    from solar_forecast.pipeline import ForecastPipeline, PipelineConfig

    config = PipelineConfig(
        target_column=args.target,
        data_path=Path(args.data) if args.data else None,
        input_dir=Path(args.input_dir),
        feature_columns=parse_csv_values(args.features),
        output_dir=Path(args.output_dir),
        sequence=build_sequence_config(args),
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


def handle_hybrid_command(args: argparse.Namespace) -> None:
    from solar_forecast.models.hybrid.experiment import HybridExperiment
    from solar_forecast.evaluation import FeatureAblationService
    from solar_forecast.jobs.training_job import TrainingService

    paths = HybridExperiment(Path(args.output_dir), args.artifact_level).run(
        Path(args.validation), Path(args.test)
    )
    print(f"Hybrid completed: {args.output_dir}")
    print(f"National metrics: {paths['national_metrics']}")


def handle_train_command(args: argparse.Namespace) -> None:
    from solar_forecast.models.hybrid.experiment import HybridExperiment
    from solar_forecast.evaluation import FeatureAblationService
    from solar_forecast.jobs.training_job import TrainingService

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


def handle_evaluate_features_command(args: argparse.Namespace) -> None:
    from solar_forecast.models.hybrid.experiment import HybridExperiment
    from solar_forecast.evaluation import FeatureAblationService
    from solar_forecast.jobs.training_job import TrainingService

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
