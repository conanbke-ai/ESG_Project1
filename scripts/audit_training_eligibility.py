"""Run the Solar Gold training-data eligibility audit without starting model training."""
from __future__ import annotations

import argparse
import csv
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from solar_forecast.preprocessing.training_eligibility import run_training_eligibility_audit


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Audit per-plant generation/KMA overlap, continuity and temporal-split "
            "coverage before freezing the model training population."
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


def _write_csv_summary(report: dict, json_path: Path) -> Path:
    csv_path = json_path.with_suffix(".csv")
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "plant_id",
        "status",
        "hard_reject_reasons",
        "target_start",
        "target_end",
        "target_rows",
        "target_expected_hourly_rows",
        "target_hourly_coverage",
        "target_longest_strict_run_hours",
        "target_gap_count_gt_1h",
        "target_max_gap_hours",
        "overlap_start",
        "overlap_end",
        "overlap_rows",
        "overlap_expected_hourly_rows",
        "overlap_hourly_coverage",
        "overlap_longest_strict_run_hours",
        "overlap_gap_count_gt_1h",
        "overlap_max_gap_hours",
        "train_overlap_rows",
        "validation_overlap_rows",
        "calibration_overlap_rows",
        "test_overlap_rows",
    ]
    with csv_path.open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.DictWriter(stream, fieldnames=fieldnames)
        writer.writeheader()
        for item in report["plants"]:
            target = item["target_continuity"]
            overlap = item["generation_weather_overlap"]
            split_rows = {
                name: item["splits"][name]["generation_weather_overlap"]["rows"]
                for name in ("train", "validation", "calibration", "test")
            }
            writer.writerow(
                {
                    "plant_id": item["plant_id"],
                    "status": item["status"],
                    "hard_reject_reasons": ";".join(item["hard_reject_reasons"]),
                    "target_start": target["start"],
                    "target_end": target["end"],
                    "target_rows": target["rows"],
                    "target_expected_hourly_rows": target["expected_hourly_rows"],
                    "target_hourly_coverage": target["hourly_coverage"],
                    "target_longest_strict_run_hours": target["longest_strict_run_hours"],
                    "target_gap_count_gt_1h": target["gap_count_gt_1h"],
                    "target_max_gap_hours": target["max_gap_hours"],
                    "overlap_start": overlap["start"],
                    "overlap_end": overlap["end"],
                    "overlap_rows": overlap["rows"],
                    "overlap_expected_hourly_rows": overlap["expected_hourly_rows"],
                    "overlap_hourly_coverage": overlap["hourly_coverage"],
                    "overlap_longest_strict_run_hours": overlap["longest_strict_run_hours"],
                    "overlap_gap_count_gt_1h": overlap["gap_count_gt_1h"],
                    "overlap_max_gap_hours": overlap["max_gap_hours"],
                    "train_overlap_rows": split_rows["train"],
                    "validation_overlap_rows": split_rows["validation"],
                    "calibration_overlap_rows": split_rows["calibration"],
                    "test_overlap_rows": split_rows["test"],
                }
            )
    return csv_path


def main() -> int:
    args = build_parser().parse_args()
    output = Path(args.output)
    report = run_training_eligibility_audit(
        Path(args.config),
        data_path=Path(args.data) if args.data else None,
        output_path=output,
        project_root=PROJECT_ROOT,
    )
    summary = _write_csv_summary(report, output)
    population = report["population"]
    print(
        "Training-data eligibility audit complete: "
        f"plants={population['plants']}, "
        f"structurally_eligible={population['candidate_plants']}, "
        f"structural_rejects={population['structurally_rejected_plants']}, "
        f"quality_eligible_rows={population['quality_eligible_target_rows']}"
    )
    print(f"fixed_start_year_used={report['fixed_start_year_used']}")
    print(f"structural_rules_applied={report['structural_rules_applied']}")
    print(f"final_training_selection_ready={report['final_training_selection_ready']}")
    print(f"manifest={output}")
    print(f"summary_csv={summary}")
    print(f"next_decision={report['next_decision']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
