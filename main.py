"""Command-line entrypoint for CNN-BiLSTM workflows."""
from __future__ import annotations

import argparse
from pathlib import Path
from typing import List, Optional

import pandas as pd

from cnn_bilstm.data_utils import SequenceConfig
from cnn_bilstm.workflows import compare_checkpoints, evaluate_and_analyze, train_and_save


def _parse_features(raw: Optional[str]) -> Optional[List[str]]:
    if not raw:
        return None
    return [col.strip() for col in raw.split(",") if col.strip()]


def _sequence_config_from_args(args: argparse.Namespace) -> SequenceConfig:
    return SequenceConfig(
        sequence_length=args.sequence_length,
        test_size=args.test_size,
        val_size=args.val_size,
        batch_size=args.batch_size,
        shuffle=not args.no_shuffle,
        num_workers=args.num_workers,
    )


def _add_sequence_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--sequence-length", type=int, default=24, help="Sequence length for sliding window datasets")
    parser.add_argument("--test-size", type=float, default=0.2, help="Proportion reserved for test split")
    parser.add_argument("--val-size", type=float, default=0.2, help="Proportion reserved for validation split")
    parser.add_argument("--batch-size", type=int, default=64, help="Batch size for dataloaders")
    parser.add_argument("--no-shuffle", action="store_true", help="Disable shuffling before train/val/test split")
    parser.add_argument("--num-workers", type=int, default=0, help="Number of workers for dataloaders")


def run_train(args: argparse.Namespace) -> None:
    df = pd.read_csv(args.data)
    seq_cfg = _sequence_config_from_args(args)
    features = _parse_features(args.features)

    artifacts = train_and_save(
        df,
        target_column=args.target,
        feature_columns=features,
        sequence_config=seq_cfg,
        n_trials=args.n_trials,
        output_dir=args.output_dir,
        use_optuna=not args.no_optuna,
        use_reinforcement=args.reinforcement,
        epochs=args.epochs,
    )

    print("Run directory:", artifacts["output_dir"])
    print("Checkpoint:", artifacts["checkpoint_path"])


def run_compare(args: argparse.Namespace) -> None:
    df = pd.read_csv(args.data)
    seq_cfg = _sequence_config_from_args(args)
    features = _parse_features(args.features)

    summary = compare_checkpoints(
        args.checkpoint_dir,
        df,
        target_column=args.target,
        feature_columns=features,
        sequence_config=seq_cfg,
    )
    if args.output:
        Path(args.output).parent.mkdir(parents=True, exist_ok=True)
        summary.to_csv(args.output, index=False)
    print(summary)


def run_analyze(args: argparse.Namespace) -> None:
    df = pd.read_csv(args.data)
    seq_cfg = _sequence_config_from_args(args)
    features = _parse_features(args.features)

    result = evaluate_and_analyze(
        args.checkpoint,
        df,
        target_column=args.target,
        feature_columns=features,
        sequence_config=seq_cfg,
        contamination=args.contamination,
        output_dir=args.output_dir,
    )
    print("Metrics:", result["metrics"])
    print("Anomalies head:\n", result["anomalies"].head())
    if result.get("output_dir"):
        print("Saved outputs to:", result["output_dir"])


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="CNN-BiLSTM workflows")
    subparsers = parser.add_subparsers(dest="command")

    train_parser = subparsers.add_parser("train", help="Run Optuna/reinforcement training and save artifacts")
    train_parser.add_argument("data", help="Path to CSV dataset")
    train_parser.add_argument("target", help="Target column name")
    train_parser.add_argument("--features", help="Comma-separated feature column names. Defaults to all but target")
    train_parser.add_argument("--output-dir", default="test/checkpoints", help="Base directory for checkpoint outputs")
    train_parser.add_argument("--n-trials", type=int, default=10, help="Number of Optuna trials")
    train_parser.add_argument("--epochs", type=int, default=50, help="Epochs for fallback training")
    train_parser.add_argument("--no-optuna", action="store_true", help="Disable Optuna search")
    train_parser.add_argument("--reinforcement", action="store_true", help="Enable reinforcement learning scheduler")
    _add_sequence_args(train_parser)
    train_parser.set_defaults(func=run_train)

    compare_parser = subparsers.add_parser("compare", help="Benchmark checkpoints against a dataset")
    compare_parser.add_argument("checkpoint_dir", help="Directory containing .pt checkpoints")
    compare_parser.add_argument("data", help="Path to CSV dataset")
    compare_parser.add_argument("target", help="Target column name")
    compare_parser.add_argument("--features", help="Comma-separated feature column names")
    compare_parser.add_argument("--output", help="Optional CSV path to save the benchmark table")
    _add_sequence_args(compare_parser)
    compare_parser.set_defaults(func=run_compare)

    analyze_parser = subparsers.add_parser("analyze", help="Reload a checkpoint, evaluate, and run anomaly detection")
    analyze_parser.add_argument("checkpoint", help="Path to a saved model checkpoint (.pt)")
    analyze_parser.add_argument("data", help="Path to CSV dataset")
    analyze_parser.add_argument("target", help="Target column name")
    analyze_parser.add_argument("--features", help="Comma-separated feature column names")
    analyze_parser.add_argument("--contamination", type=float, default=0.05, help="IsolationForest contamination proportion")
    analyze_parser.add_argument("--output-dir", help="Base directory for analysis artifacts")
    _add_sequence_args(analyze_parser)
    analyze_parser.set_defaults(func=run_analyze)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    if not getattr(args, "command", None):
        parser.print_help()
        return

    args.func(args)


if __name__ == "__main__":
    main()
