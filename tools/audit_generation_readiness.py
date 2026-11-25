"""Audit dated meter and capacity evidence for saved generation source series."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from solar_forecast.collectors.plant_identity import resolve_generation_identity
from solar_forecast.collectors.plant_metadata import PlantMetadataCatalog
from solar_forecast.datasets.archive_standardizer import (
    COMPANY_DIRECTORIES,
    HistoricalGenerationStandardizationService,
)
from solar_forecast.infrastructure.artifact_store import sha256_file, write_json_atomic


_FIELDS = ["timestamp", "company", "plant_id", "plant", "unit", "capacity_mw"]
_REVIEW_CONTRACT = "solar-generation-asset-evidence.v1"


def _local_hour(value: str) -> pd.Timestamp:
    stamp = pd.Timestamp(value)
    if pd.isna(stamp) or stamp.tzinfo is not None or stamp != stamp.floor("h"):
        raise ValueError("Evidence bounds must be local KST hour labels")
    return stamp


def _read_review(path: Path | None) -> dict:
    """Verify review references against local source bytes and validate intervals."""
    if path is None:
        return {"evidence": {}, "meter_bindings": {}, "capacity_periods": {}}
    review = json.loads(path.read_text(encoding="utf-8-sig"))
    if review.get("contract") != _REVIEW_CONTRACT:
        raise ValueError(f"Expected review contract {_REVIEW_CONTRACT}")
    result = {"evidence": {}, "meter_bindings": {}, "capacity_periods": {}}
    for evidence_id, entry in review.get("evidence", {}).items():
        evidence_path = path.parent / entry["path"]
        digest = sha256_file(evidence_path)
        if digest != entry["sha256"]:
            raise ValueError(f"Evidence SHA-256 mismatch: {evidence_id}")
        result["evidence"][evidence_id] = {"path": str(evidence_path), "sha256": digest}
    for kind in ("meter_bindings", "capacity_periods"):
        for entry in review.get(kind, []):
            series_id = entry["source_series_id"].strip()
            start = _local_hour(entry["start_kst"])
            end = _local_hour(entry["end_exclusive_kst"])
            if not series_id or start >= end or entry["evidence_id"] not in result["evidence"]:
                raise ValueError(f"Invalid {kind} entry for {series_id!r}")
            if kind == "meter_bindings":
                field = "meter_point_id"
                value = entry[field].strip()
                if not value:
                    raise ValueError(f"Empty physical meter point for {series_id}")
            else:
                field = "capacity_mw"
                value = entry[field]
                if type(value) not in (int, float) or not np.isfinite(value) or value <= 0:
                    raise ValueError(f"Evidence capacity must be a positive finite MW value: {series_id}")
            result[kind].setdefault(series_id, []).append({
                "source_series_id": series_id,
                "start_kst": start.isoformat(),
                "end_exclusive_kst": end.isoformat(),
                "evidence_id": entry["evidence_id"],
                field: value,
                "_start": start,
                "_end": end,
            })
        for series_id, entries in result[kind].items():
            entries.sort(key=lambda item: item["_start"])
            if any(left["_end"] > right["_start"] for left, right in zip(entries, entries[1:])):
                raise ValueError(f"Overlapping {kind} intervals for {series_id}")
    return result


def _read_input(path: Path, company: str | None) -> pd.DataFrame:
    """Use the archived normalizers or the saved standardized input contract."""
    if company is None:
        input_fields = {*_FIELDS, "energy_source"}
        frame = pd.read_csv(path, usecols=lambda column: column in input_fields, dtype={"plant_id": str})
        required = set(_FIELDS) - {"unit", "capacity_mw"}
        if missing := required - set(frame):
            raise ValueError(f"Missing standardized columns in {path}: {sorted(missing)}")
        frame = resolve_generation_identity(frame)
    else:
        archive = HistoricalGenerationStandardizationService(
            path.parent, path.parent, metadata=PlantMetadataCatalog(())
        )
        frame = archive._normalize(company, path)
    timestamps = pd.to_datetime(frame["timestamp"], format="mixed", errors="raise")
    if (
        not pd.api.types.is_datetime64_any_dtype(timestamps.dtype)
        or timestamps.dt.tz is not None
        or not timestamps.eq(timestamps.dt.floor("h")).all()
    ):
        raise ValueError(f"Source timestamps must be complete local KST hour labels: {path}")
    frame["timestamp"] = timestamps
    for field in ("plant_id", "plant", "company", "unit"):
        labels = frame.get(field, pd.Series(pd.NA, index=frame.index)).astype("string").str.strip()
        frame[field] = labels.mask(labels.isin(["", "nan", "None", "<NA>"]))
    frame["capacity_mw"] = pd.to_numeric(frame.get("capacity_mw", np.nan), errors="coerce")
    return frame[_FIELDS]


def _capacity_observations(rows: pd.DataFrame) -> list[dict]:
    """List supplied MW values by calendar year and date bounds of occurrence."""
    records = []
    for year, period in rows.groupby(rows["timestamp"].dt.year):
        values = period["capacity_mw"]
        known = period.loc[values.gt(0) & np.isfinite(values)]
        records.append({
            "year": int(year),
            "source_rows": len(period),
            "missing_capacity_rows": int(values.isna().sum()),
            "invalid_capacity_rows": int(values.notna().sum() - len(known)),
            "observed_values": [
                {
                    "capacity_mw": float(capacity),
                    "rows": len(samples),
                    "first_observed_hour": samples["timestamp"].min().isoformat(),
                    "last_observed_hour": samples["timestamp"].max().isoformat(),
                }
                for capacity, samples in known.groupby("capacity_mw", sort=True)
            ],
        })
    return records


def _interval_coverage(rows: pd.DataFrame, entries: list[dict], *, capacity: bool) -> dict:
    """Show where reviewed records apply to this source series and expose mismatches."""
    covered = pd.Series(False, index=rows.index)
    conflicts = pd.Series(False, index=rows.index)
    periods = []
    for entry in entries:
        scope = rows["timestamp"].ge(entry["_start"]) & rows["timestamp"].lt(entry["_end"])
        covered |= scope
        if capacity:
            values = rows["capacity_mw"]
            conflicts |= scope & values.notna() & ~np.isclose(
                values, entry["capacity_mw"], rtol=0, atol=1e-9, equal_nan=False
            )
        periods.append({
            **{key: value for key, value in entry.items() if not key.startswith("_")},
            "observed_hours": int(rows.loc[scope, "timestamp"].nunique()),
        })
    report = {
        "covered_hours": int(rows.loc[covered, "timestamp"].nunique()),
        "uncovered_hours": int(rows.loc[~covered, "timestamp"].nunique()),
        "periods": periods,
    }
    if capacity:
        report["reported_capacity_conflict_hours"] = int(rows.loc[conflicts, "timestamp"].nunique())
    return report


def _overlapping_meter_bindings(bindings: dict[str, list[dict]]) -> list[dict]:
    """Expose concurrent aliases that could count the same physical meter twice."""
    by_point = {}
    for series_id, entries in bindings.items():
        for entry in entries:
            by_point.setdefault(entry["meter_point_id"], []).append((series_id, entry))
    findings = []
    for point, entries in sorted(by_point.items()):
        entries.sort(key=lambda pair: pair[1]["_start"])
        for index, (left_id, left) in enumerate(entries):
            for right_id, right in entries[index + 1:]:
                if right["_start"] >= left["_end"]:
                    break
                if left_id != right_id:
                    findings.append({
                        "meter_point_id": point,
                        "source_series_ids": sorted([left_id, right_id]),
                        "start_kst": right["_start"].isoformat(),
                        "end_exclusive_kst": min(left["_end"], right["_end"]).isoformat(),
                    })
    return findings


def audit_generation_readiness(
    inputs: tuple[tuple[str | None, Path], ...], review_path: Path | None = None,
) -> dict:
    """Read only the listed archives/Silver files and report date-scoped evidence."""
    if not inputs or len({path.resolve() for _, path in inputs}) != len(inputs):
        raise ValueError("Provide one or more distinct input paths")
    review = _read_review(review_path)
    source_files = []
    batches = []
    for order, (company, path) in enumerate(inputs):
        frame = _read_input(path, company)
        source_files.append({
            "order": order,
            "path": str(path),
            "sha256": sha256_file(path),
            "normalized_rows": len(frame),
            "missing_source_identity_rows": int(frame[["company", "plant_id", "plant"]].isna().any(axis=1).sum()),
            "source_identity_resolutions": frame.attrs.get("plant_identity_resolutions", []),
        })
        batches.append(frame.dropna(subset=["company", "plant_id", "plant"]).assign(source_order=order))
    data = pd.concat(batches, ignore_index=True)
    series = []
    for series_id, rows in data.groupby("plant_id", sort=True):
        meter = _interval_coverage(rows, review["meter_bindings"].get(series_id, []), capacity=False)
        capacity = _interval_coverage(rows, review["capacity_periods"].get(series_id, []), capacity=True)
        series.append({
            "source_series_id": series_id,
            "company_labels": sorted(rows["company"].unique()),
            "plant_labels": sorted(rows["plant"].unique()),
            "generator_or_unit_labels": sorted(rows["unit"].dropna().unique()),
            "missing_generator_or_unit_label_rows": int(rows["unit"].isna().sum()),
            "observed_hours": int(rows["timestamp"].nunique()),
            "meter_point_evidence": meter,
            "capacity_period_evidence": capacity,
            "observed_capacity_by_source": [
                {"source_order": int(order), "years": _capacity_observations(group)}
                for order, group in rows.groupby("source_order", sort=True)
            ],
        })
    reviewed_ids = set(review["meter_bindings"]) | set(review["capacity_periods"])
    return {
        "contract": "solar-generation-readiness.v1",
        "timestamp_timezone": "Asia/Seoul",
        "scope": "Existing normalized source hours. Observed annual capacity bounds are occurrence dates, never inferred effective intervals.",
        "meter_identity": "Generator labels and meter classifications are source labels. Only separately reviewed physical-meter intervals cover a series.",
        "capacity_conflict_tolerance_mw": 1e-9,
        "review_file": (
            {"path": str(review_path), "sha256": sha256_file(review_path)}
            if review_path is not None else None
        ),
        "reviewed_evidence": review["evidence"],
        "overlapping_meter_bindings": _overlapping_meter_bindings(review["meter_bindings"]),
        "review_series_not_seen": sorted(reviewed_ids - {entry["source_series_id"] for entry in series}),
        "source_files": source_files,
        "source_series": series,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--archive-file", action="append", default=[], metavar="COMPANY=PATH")
    parser.add_argument("--partition", type=Path, action="append", default=[])
    parser.add_argument("--review", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    inputs = []
    for value in args.archive_file:
        company, separator, path = value.partition("=")
        if company not in COMPANY_DIRECTORIES or not separator or not path:
            parser.error("--archive-file requires an existing archive company and path")
        inputs.append((company, Path(path)))
    inputs.extend((None, path) for path in args.partition)
    if args.output.suffix.lower() != ".json" or args.output.exists():
        parser.error("--output must be a new .json report path")
    report = audit_generation_readiness(tuple(inputs), args.review)
    write_json_atomic(args.output, report)
    print(args.output)


if __name__ == "__main__":
    main()
