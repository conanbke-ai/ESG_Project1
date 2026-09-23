"""Audit plant/month/season split coverage without importing model frameworks."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from solar_forecast.evaluation.split_audit import run_split_audit


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=Path("config/experiments/optimized.json"))
    parser.add_argument("--data", type=Path, help="Override the observed Gold CSV/CSV.GZ or partition directory")
    parser.add_argument("--output", type=Path, default=Path("artifacts/data_audits/temporal_split.json"))
    args = parser.parse_args()
    report = run_split_audit(args.config, data_path=args.data, output_path=args.output)
    print(json.dumps({"report": str(args.output), "audits": [
        {"horizons_hours": item["horizons_hours"], "boundaries": item["boundaries"],
         "population": item["population"], "splits": item["splits"], "warnings": item["warnings"]}
        for item in report["audits"]
    ]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
