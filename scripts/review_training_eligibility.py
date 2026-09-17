"""Summarize the real Solar training-eligibility evidence before freezing thresholds."""
from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
import statistics


CANDIDATE_STATUS = "CANDIDATE_REQUIRES_THRESHOLD_REVIEW"
SPLITS = ("train", "validation", "calibration", "test")


def _quantile(values: list[float], q: float) -> float:
    ordered = sorted(values)
    if not ordered:
        return 0.0
    if len(ordered) == 1:
        return ordered[0]
    position = (len(ordered) - 1) * q
    lower = int(position)
    upper = min(lower + 1, len(ordered) - 1)
    weight = position - lower
    return ordered[lower] * (1 - weight) + ordered[upper] * weight


def _distribution(values: list[float]) -> dict[str, float]:
    return {
        "min": min(values),
        "p10": _quantile(values, 0.10),
        "p25": _quantile(values, 0.25),
        "median": statistics.median(values),
        "p75": _quantile(values, 0.75),
        "p90": _quantile(values, 0.90),
        "max": max(values),
    }


def _candidate_row(item: dict) -> dict[str, object]:
    overlap = item["generation_weather_overlap"]
    split_rows = {
        name: int(item["splits"][name]["generation_weather_overlap"]["rows"])
        for name in SPLITS
    }
    total_split_rows = sum(split_rows.values())
    weather_coverages = [
        float(value["coverage"])
        for value in item.get("weather_column_coverage", {}).values()
    ]
    return {
        "plant_id": item["plant_id"],
        "overlap_start": overlap["start"],
        "overlap_end": overlap["end"],
        "overlap_rows": int(overlap["rows"]),
        "overlap_hourly_coverage": float(overlap["hourly_coverage"]),
        "overlap_longest_strict_run_hours": int(overlap["longest_strict_run_hours"]),
        "overlap_gap_count_gt_1h": int(overlap["gap_count_gt_1h"]),
        "overlap_max_gap_hours": float(overlap["max_gap_hours"] or 0.0),
        "weather_coverage_min": min(weather_coverages) if weather_coverages else 0.0,
        "weather_coverage_median": statistics.median(weather_coverages) if weather_coverages else 0.0,
        **{f"{name}_rows": split_rows[name] for name in SPLITS},
        **{
            f"{name}_fraction": (split_rows[name] / total_split_rows if total_split_rows else 0.0)
            for name in SPLITS
        },
    }


def _write_csv(rows: list[dict[str, object]], destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Summarize the 15-ish structurally valid plant candidates before threshold freeze."
    )
    parser.add_argument(
        "--manifest",
        default="artifacts/evaluation/training_data_eligibility.json",
    )
    parser.add_argument(
        "--output",
        default="artifacts/evaluation/training_data_eligibility_review.csv",
    )
    args = parser.parse_args()

    manifest = Path(args.manifest)
    report = json.loads(manifest.read_text(encoding="utf-8"))
    candidates = [
        _candidate_row(item)
        for item in report["plants"]
        if item["status"] == CANDIDATE_STATUS
    ]
    if not candidates:
        raise SystemExit("No threshold-review candidates found in the eligibility manifest")

    candidates.sort(
        key=lambda row: (
            row["test_rows"],
            row["calibration_rows"],
            row["validation_rows"],
            row["train_rows"],
            row["overlap_hourly_coverage"],
        )
    )
    output = Path(args.output)
    _write_csv(candidates, output)

    metric_names = [
        "overlap_rows",
        "overlap_hourly_coverage",
        "overlap_longest_strict_run_hours",
        "overlap_gap_count_gt_1h",
        "overlap_max_gap_hours",
        "weather_coverage_min",
        "weather_coverage_median",
        "train_rows",
        "validation_rows",
        "calibration_rows",
        "test_rows",
    ]
    distributions = {
        name: _distribution([float(row[name]) for row in candidates])
        for name in metric_names
    }

    print(f"candidate_plants={len(candidates)}")
    print("\n=== DISTRIBUTIONS ===")
    for name in metric_names:
        values = distributions[name]
        print(
            f"{name}: min={values['min']:.6g} p10={values['p10']:.6g} "
            f"p25={values['p25']:.6g} median={values['median']:.6g} "
            f"p75={values['p75']:.6g} p90={values['p90']:.6g} max={values['max']:.6g}"
        )

    print("\n=== CANDIDATES (weakest split volume first) ===")
    header = (
        "plant_id | overlap_cov | longest_h | weather_min | "
        "train | val | cal | test | split_fraction(T/V/C/Test)"
    )
    print(header)
    for row in candidates:
        print(
            f"{row['plant_id']} | {row['overlap_hourly_coverage']:.4f} | "
            f"{row['overlap_longest_strict_run_hours']} | {row['weather_coverage_min']:.4f} | "
            f"{row['train_rows']} | {row['validation_rows']} | {row['calibration_rows']} | {row['test_rows']} | "
            f"{row['train_fraction']:.3f}/{row['validation_fraction']:.3f}/"
            f"{row['calibration_fraction']:.3f}/{row['test_fraction']:.3f}"
        )

    print("\n=== DATA-DRIVEN REVIEW FLAGS (NOT FINAL REJECTION RULES) ===")
    coverage_q25 = distributions["overlap_hourly_coverage"]["p25"]
    continuity_q25 = distributions["overlap_longest_strict_run_hours"]["p25"]
    weather_q25 = distributions["weather_coverage_min"]["p25"]
    split_q25 = {
        name: distributions[f"{name}_rows"]["p25"]
        for name in SPLITS
    }
    for row in candidates:
        flags: list[str] = []
        if row["overlap_hourly_coverage"] < coverage_q25:
            flags.append("low_overlap_coverage_vs_candidates")
        if row["overlap_longest_strict_run_hours"] < continuity_q25:
            flags.append("short_continuity_vs_candidates")
        if row["weather_coverage_min"] < weather_q25:
            flags.append("low_weather_coverage_vs_candidates")
        for name in SPLITS:
            if row[f"{name}_rows"] < split_q25[name]:
                flags.append(f"low_{name}_volume_vs_candidates")
        print(f"{row['plant_id']}: {','.join(flags) if flags else 'none'}")

    print(f"\nreview_csv={output}")
    print("final_thresholds_applied=False")
    print("Paste the DISTRIBUTIONS + CANDIDATES output into ChatGPT to freeze the final admission policy.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
