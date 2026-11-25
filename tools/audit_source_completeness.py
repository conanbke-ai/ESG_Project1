"""Audit retained KOEN, EWP, and KOSPO source coverage without fetching or rewriting files."""

from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from solar_forecast.datasets.source_completeness import audit_source_completeness
from solar_forecast.infrastructure.artifact_store import write_json_atomic


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--archive-root", type=Path, default=Path("file/solar_data_file"))
    parser.add_argument("--raw-root", type=Path, default=Path("file/raw"))
    parser.add_argument("--collection-index", type=Path, action="append", default=[])
    parser.add_argument("--as-of", type=date.fromisoformat, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    output = args.output.resolve()
    source_roots = (args.archive_root.resolve(), args.raw_root.resolve())
    if any(output.is_relative_to(root) for root in source_roots):
        parser.error("The report must be written outside the source roots.")
    if output in {path.resolve() for path in args.collection_index}:
        parser.error("The report cannot overwrite a collection index.")
    report = audit_source_completeness(
        args.archive_root, args.raw_root, args.as_of, collection_indexes=tuple(args.collection_index),
    )
    write_json_atomic(args.output, report)
    print(args.output)


if __name__ == "__main__":
    main()
