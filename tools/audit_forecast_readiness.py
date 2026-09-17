"""학습 없이 실제 예측 표본·연속 입력창·Train 피처를 점검."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from solar_forecast.evaluation.forecast_readiness import run_forecast_readiness


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=Path("config/experiments/observed_calendar_candidate.json"))
    parser.add_argument("--data", type=Path, help="Observed Gold CSV, CSV.GZ, or partition directory")
    parser.add_argument("--output", type=Path, default=Path("artifacts/data_audits/forecast_readiness.json"))
    args = parser.parse_args()
    report = run_forecast_readiness(args.config, data_path=args.data, output_path=args.output)
    print(json.dumps({"report": str(args.output), "coverage_gate_passed": report["coverage_gate_passed"],
                      "prediction_or_training_performed": False,
                      "audits": [{"horizon_hours": audit["horizon_hours"], "common_coverage": audit["common_coverage"],
                                  "candidates": [{"model": item["model"], "candidate_id": item["candidate_id"],
                                                   "samples": {name: split["usable_samples"] for name, split in item["splits"].items()},
                                                   "warnings": item["warnings"]} for item in audit["candidates"]]} for audit in report["audits"]]},
                     ensure_ascii=False, indent=2))
    if not report["coverage_gate_passed"]:
        raise SystemExit(2)


if __name__ == "__main__":
    main()
