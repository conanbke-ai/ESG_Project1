"""Collection commands: CLI input translation and command dispatch."""
from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
from solar_forecast.cli.arguments import parse_csv_values


def handle_collect_command(args: argparse.Namespace) -> None:
    from solar_forecast.collectors import CollectionConfig, CollectionService

    config = CollectionConfig(
        start_date=date.fromisoformat(args.start_date),
        end_date=date.fromisoformat(args.end_date) if args.end_date else date.today(),
        sources=parse_csv_values(args.sources) or [],
        output_dir=Path(args.output_dir),
        standardized_output_dir=Path(args.standardized_output_dir),
        overwrite=args.overwrite,
        komipo_station_codes=parse_csv_values(args.komipo_station_codes) or [],
        api_max_calls=args.api_max_calls,
        station_ids=parse_csv_values(args.station_ids) or [],
        existing_weather_dir=Path(args.weather_root),
        kma_mode=args.kma_mode,
        download_date=date.fromisoformat(args.download_date) if args.download_date else date.today(),
    )
    results = CollectionService(config).run()
    for result in results:
        print(f"{result.source}: {result.status} ({len(result.files)} files, {result.rows} rows) {result.message}")
    if any(result.status in {"failed", "configuration_required", "unsupported"} for result in results):
        raise SystemExit(1)
