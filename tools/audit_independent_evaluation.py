"""Audit the 2026 holdout gate and verify the saved development comparison."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from solar_forecast.evaluation.independent_evaluation import run_independent_evaluation_audit


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=Path("config/experiments/optimized.json"))
    parser.add_argument("--development-run", type=Path, required=True, help="Saved 2025 benchmark run directory")
    parser.add_argument("--admission", type=Path, required=True, help="P1 Gold candidate manifest")
    parser.add_argument("--exposure-review", type=Path, help="Hashed, completed development-use review for the candidate")
    parser.add_argument("--output", type=Path, required=True, help="New report file")
    args = parser.parse_args()
    report = run_independent_evaluation_audit(config_path=args.config, development_run=args.development_run,
                                               admission_path=args.admission, exposure_path=args.exposure_review,
                                               output_path=args.output)
    print(json.dumps({"report": str(args.output), "status": report["status"],
                      "candidate_target_rows": report["holdout"]["candidate_target_rows"],
                      "independent_scoring_performed": report["independent_scoring_performed"]}, indent=2))


if __name__ == "__main__":
    main()
