"""Run the Solar Gold training-data eligibility audit without starting model training."""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from solar_forecast.evaluation.training_eligibility import run_training_eligibility_audit


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Audit per-plant generation/KMA overlap, continuity and temporal-split "
            "coverage before freezing final training admission thresholds."
        )
    )
    parser.add_argument(
        "--config",
        default="config/experiments/optimized.json",
        help="Experiment config that owns the current temporal split contract.",
    )
    parser.add_argument(
        "--data",
        help="Optional Gold CSV/CSV.GZ or partition directory override.",
    )
    parser.add_argument(
        "--output",
        default="artifacts/evaluation/training_data_eligibility.json",
        help="JSON evidence manifest written outside the source dataset.",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    report = run_training_eligibility_audit(
        Path(args.config),
        data_path=Path(args.data) if args.data else None,
        output_path=Path(args.output) if args.output else None,
        project_root=PROJECT_ROOT,
    )
    population = report["population"]
    print(
        "Training-data eligibility audit complete: "
        f"plants={population['plants']}, "
        f"candidates={population['candidate_plants']}, "
        f"structural_rejects={population['structurally_rejected_plants']}, "
        f"quality_eligible_rows={population['quality_eligible_target_rows']}"
    )
    print(f"fixed_start_year_used={report['fixed_start_year_used']}")
    print(f"final_thresholds_applied={report['final_thresholds_applied']}")
    print(f"final_training_selection_ready={report['final_training_selection_ready']}")
    print(f"manifest={args.output}")
    print(f"next_decision={report['next_decision']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
