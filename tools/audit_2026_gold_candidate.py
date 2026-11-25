"""Audit retained 2026 plant evidence and write a separate Gold candidate manifest."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from solar_forecast.datasets.phase2_admission import (
    BASELINE_MANIFEST,
    BASELINE_VERIFICATION,
    CONTRACT,
    DATA_VERSION,
    HORIZONS,
    WEATHER_STAGING,
    GenerationSource,
    audit_2026_gold_candidate,
)
from solar_forecast.infrastructure.artifact_store import write_json_atomic


DEFAULT_OUTPUT = Path("docs/evidence/data_admission_phase2/2026_gold_candidate.json")


def _assignment(value: str) -> tuple[str, Path]:
    key, sep, filename = value.partition("=")
    if not sep or not key or not filename:
        raise argparse.ArgumentTypeError("Expected KEY=PATH")
    return key, Path(filename)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd(), help=r"Data checkout, e.g. C:\ESG_Project1")
    parser.add_argument("--generation", action="append", type=_assignment, default=[],
                        metavar="COMPANY=PATH", help="Additional retained raw or collector Silver input (repeatable)")
    parser.add_argument("--source-audit", type=Path, help="Reuse an existing source completeness report and its source hashes")
    parser.add_argument("--asos", type=Path, help="2026 standard-format ASOS observation CSV")
    parser.add_argument("--weather-staging-manifest", type=Path, default=WEATHER_STAGING)
    parser.add_argument("--forecast", action="append", type=_assignment, default=[], metavar="HOURS=PATH",
                        help="Per-horizon archive override (repeatable)")
    parser.add_argument("--observation-availability", type=Path, help="CSV with actual per-hour availability provenance")
    parser.add_argument("--plant-review", type=Path, help="JSON of source-backed meter and interval reviews by plant")
    parser.add_argument("--baseline-manifest", type=Path, default=BASELINE_MANIFEST)
    parser.add_argument("--gold-verification", type=Path, default=BASELINE_VERIFICATION)
    parser.add_argument("--candidate-gold-manifest", type=Path, help=f"Existing Gold manifest in a separate {DATA_VERSION}/ directory")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    root = args.root.resolve()
    if any(company not in {"koen", "ewp", "kospo"} for company, _ in args.generation):
        parser.error("--generation accepts koen, ewp or kospo")
    horizon_names = [horizon for horizon, _ in args.forecast]
    if len(set(horizon_names)) != len(horizon_names) or set(horizon_names) - set(map(str, HORIZONS)):
        parser.error("--forecast accepts one override each for 1, 24 or 72 hours")
    output = (root / args.output).resolve()
    input_paths = [
        args.baseline_manifest, args.gold_verification, args.weather_staging_manifest, args.candidate_gold_manifest,
        args.plant_review, args.observation_availability, args.asos, args.source_audit,
        *(path for _, path in args.generation), *(path for _, path in args.forecast),
    ]
    if output.suffix != ".json" or any(output == (root / path).resolve() for path in input_paths if path):
        parser.error("--output must be a separate JSON report")
    if output.exists():
        previous = json.loads(output.read_text(encoding="utf-8"))
        if previous.get("contract") != CONTRACT or previous.get("data_version") != DATA_VERSION:
            parser.error("--output would replace a different artifact")

    report = audit_2026_gold_candidate(
        root, generation_sources=tuple(GenerationSource(company, path) for company, path in args.generation),
        availability_path=args.observation_availability, review_path=args.plant_review,
        forecast_paths={int(horizon): path for horizon, path in args.forecast},
        weather_csv=args.asos, weather_staging_manifest=args.weather_staging_manifest,
        baseline_manifest=args.baseline_manifest, verification=args.gold_verification,
        candidate_manifest=args.candidate_gold_manifest, source_audit=args.source_audit,
    )
    inputs = [report["inputs"], report["baseline_gold"], report["candidate"].get("versioned_gold_manifest")]
    while inputs:
        item = inputs.pop()
        if isinstance(item, list):
            inputs.extend(item)
        elif isinstance(item, dict):
            if isinstance(item.get("path"), str) and output == (root / item["path"]).resolve():
                parser.error("--output overlaps existing data or provenance evidence")
            inputs.extend(item.values())
    write_json_atomic(output, report)
    print(json.dumps({
        "manifest": str(output), "baseline_gold": report["baseline_gold"]["status"],
        "generation_2026_rows": report["generation_2026"]["parsed_rows"],
        "candidate_status": report["candidate"]["status"],
        "admitted_2026_rows": report["candidate"]["admitted_2026_rows"],
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
