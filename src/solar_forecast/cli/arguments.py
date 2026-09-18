"""Arguments: CLI input translation and command dispatch."""
from __future__ import annotations

import argparse


def parse_csv_values(value: str | None) -> list[str] | None:
    return [item.strip() for item in value.split(",") if item.strip()] if value else None


def build_sequence_config(args: argparse.Namespace):
    from solar_forecast.models.cnn_bilstm.sequence_config import SequenceConfig

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


def add_sequence_arguments(parser: argparse.ArgumentParser) -> None:
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
