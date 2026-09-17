"""Freeze Solar training population from eligibility evidence and explicit admission policy."""
from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from solar_forecast.evaluation.training_admission import classify_training_population
from solar_forecast.infrastructure.artifact_store import sha256_file, write_json_atomic


def _write_csv(report: dict, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = [
        "plant_id",
        "status",
        "reasons",
        "overlap_hourly_coverage",
        "longest_strict_run_hours",
        "weather_column_coverage_min",
        "train_rows",
        "validation_rows",
        "calibration_rows",
        "test_rows",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader()
        for decision in report["decisions"]:
            metrics = decision["metrics"]
            writer.writerow(
                {
                    "plant_id": decision["plant_id"],
                    "status": decision["status"],
                    "reasons": ";".join(decision["reasons"]),
                    "overlap_hourly_coverage": metrics["overlap_hourly_coverage"],
                    "longest_strict_run_hours": metrics["longest_strict_run_hours"],
                    "weather_column_coverage_min": metrics["weather_column_coverage_min"],
                    "train_rows": metrics["train_rows"],
                    "validation_rows": metrics["validation_rows"],
                    "calibration_rows": metrics["calibration_rows"],
                    "test_rows": metrics["test_rows"],
                }
            )
    return path


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Classify the Gold plant population as ADMITTED/HOLD/REJECT without training."
    )
    parser.add_argument(
        "--eligibility",
        default="artifacts/evaluation/training_data_eligibility.json",
    )
    parser.add_argument(
        "--policy",
        default="config/training_admission.json",
    )
    parser.add_argument(
        "--output",
        default="artifacts/evaluation/training_population.json",
    )
    args = parser.parse_args()

    eligibility_path = Path(args.eligibility)
    policy_path = Path(args.policy)
    output_path = Path(args.output)
    eligibility = json.loads(eligibility_path.read_text(encoding="utf-8"))
    policy = json.loads(policy_path.read_text(encoding="utf-8"))
    report = classify_training_population(eligibility, policy)
    report["eligibility_manifest"] = str(eligibility_path)
    report["eligibility_manifest_sha256"] = sha256_file(eligibility_path)
    report["admission_policy"] = str(policy_path)
    report["admission_policy_sha256"] = sha256_file(policy_path)
    write_json_atomic(output_path, report)
    csv_path = _write_csv(report, output_path.with_suffix(".csv"))

    counts = report["counts"]
    print(
        "Training population classification: "
        f"ADMITTED={counts['admitted']} HOLD={counts['hold']} REJECT={counts['rejected']}"
    )
    print(f"policy_status={report['policy_status']}")
    print(f"final_training_selection_ready={report['final_training_selection_ready']}")
    if report["selection_blockers"]:
        print("selection_blockers=" + ",".join(report["selection_blockers"]))
    print(f"manifest={output_path}")
    print(f"summary_csv={csv_path}")
    return 0 if report["final_training_selection_ready"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
