"""Verification commands: CLI input translation and command dispatch."""
from __future__ import annotations

import argparse
from pathlib import Path
from solar_forecast.config_loader import PROJECT_ROOT
from solar_forecast.cli.arguments import parse_csv_values


def handle_verify_e2e_command(args: argparse.Namespace) -> None:
    from solar_forecast.jobs.verification_job import (
        PipelineVerificationService,
        VerificationConfig,
        build_collection_config,
    )

    if args.collect and not args.start_date:
        raise SystemExit("--start-date is required when --collect is used")

    collection_config = None
    if args.collect:
        collection_config = build_collection_config(
            start_date=args.start_date,
            end_date=args.end_date,
            sources=args.sources,
            output_dir=args.collection_output_dir,
            standardized_output_dir=args.collection_standardized_output_dir,
            overwrite=args.overwrite,
            download_date=args.download_date,
            komipo_station_codes=tuple(parse_csv_values(args.komipo_station_codes) or []),
            api_max_calls=args.api_max_calls,
            station_ids=tuple(parse_csv_values(args.station_ids) or []),
            kma_mode=args.kma_mode,
            weather_root=args.weather_root,
        )

    result = PipelineVerificationService(
        VerificationConfig(
            project_root=PROJECT_ROOT,
            report_root=Path(args.report_root),
            collect=args.collect,
            collection_config=collection_config,
            collection_manifest=Path(args.collection_manifest)
            if args.collection_manifest
            else None,
            allow_collection_failures=args.allow_collection_failures,
            prepare_data=not args.skip_prepare,
            input_root=Path(args.input_root),
            weather_root=Path(args.weather_root),
            merged_source=Path(args.merged_source),
            standardized_output_dir=Path(args.output_dir),
            collected_generation_dir=(
                None
                if args.no_collected_downloads
                else Path(args.collected_generation_dir)
            ),
            train_models=tuple(parse_csv_values(args.models) or []),
            smoke=not args.full_train,
            no_optuna=args.no_optuna,
            max_trials=args.max_trials,
            optimizer_timeout_seconds=args.optimizer_timeout_seconds,
            build_dashboard=not args.skip_dashboard,
            dashboard_output_dir=Path(args.dashboard_output_dir),
        )
    ).run()
    print(f"E2E verification: {result.status}")
    print(f"Report: {result.report_path}")
    for step in result.steps:
        print(f"- {step['name']}: {step['status']} ({step['duration_seconds']}s)")
