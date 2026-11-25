"""Streaming verification of existing Gold exports; never build or repair data.

Only a private temporary SQLite index is written. It bounds key/cardinality
checking memory independently of the dataset size and is discarded on exit.
"""
from __future__ import annotations

from collections import Counter
import csv
from datetime import datetime
import gzip
from hashlib import sha256
import json
import math
import os
from pathlib import Path
import sqlite3
from tempfile import TemporaryDirectory

from solar_forecast.datasets.repository import DatasetRepository
from solar_forecast.infrastructure.artifact_store import sha256_file


GOLD_AUDIT_CONTRACT = "solar-gold-consistency-audit.v1"
MANIFEST_CONTRACT = "solar-model-ready-manifest.v1"


def _resolve(root: Path, value: str) -> Path:
    if "://" in value:
        raise ValueError("Gold audit accepts project-contained filesystem paths, not URLs")
    path = Path(value)
    path = Path(os.path.abspath(path if path.is_absolute() else root / path))
    if not path.is_relative_to(root):
        raise ValueError(f"Gold audit input must stay within the project: {value}")
    path = path.resolve()
    if not path.is_relative_to(root):
        raise ValueError(f"Gold audit input must stay within the project: {value}")
    return path


def _scan_file(path, representation, db, features, batch_rows):
    """Hash every CSV field, with no numeric rounding, filling or row sampling."""
    opener = gzip.open if path.name.lower().endswith(".gz") else open
    rows, duplicates, start, end = 0, 0, None, None
    invalid_eligible_targets = 0
    eligible, energy, numeric_missing, numeric_invalid = Counter(), Counter(), Counter(), Counter()
    numeric_nonfinite = Counter()
    partition_mismatches = 0
    batch = []

    def flush():
        nonlocal duplicates
        before = db.total_changes
        db.executemany("INSERT OR IGNORE INTO cells VALUES (?, ?, ?, ?, ?, ?)", batch)
        duplicates += len(batch) - (db.total_changes - before)
        batch.clear()

    with opener(path, "rt", encoding="utf-8-sig", newline="") as stream:
        reader = csv.reader(stream, strict=True)
        header = next(reader, [])
        required = {"plant_id", "timestamp", "company", "energy_source", "quality_train_eligible"}
        missing = required.union(features).difference(header)
        if len(header) != len(set(header)) or missing:
            raise ValueError(f"Invalid Gold header in {path}: duplicate columns or missing {sorted(missing)}")
        idx = {name: header.index(name) for name in required}
        numeric_indices = [(name, header.index(name)) for name in features]
        target_index = header.index(features[-1])
        for line, row in enumerate(reader, 2):
            if len(row) != len(header) or not row[idx["plant_id"]].strip():
                raise ValueError(f"Invalid Gold row width/plant_id at {path}:{line}")
            try:
                stamp = datetime.fromisoformat(row[idx["timestamp"]])
            except ValueError as exc:
                raise ValueError(f"Unparseable Gold timestamp at {path}:{line}") from exc
            if stamp.tzinfo is not None or (stamp.minute, stamp.second, stamp.microsecond) != (0, 0, 0):
                raise ValueError(f"Gold requires naive local hourly timestamps at {path}:{line}")
            key_time = stamp.isoformat()
            start = min(start, key_time) if start else key_time
            end = max(end, key_time) if end else key_time
            flag = row[idx["quality_train_eligible"]].strip().lower()
            if flag not in {"true", "1", "yes", "false", "0", "no"}:
                raise ValueError(f"Invalid quality eligibility flag at {path}:{line}")
            source = row[idx["energy_source"]]
            energy[source] += 1
            if flag in {"true", "1", "yes"}:
                eligible[source] += 1
                try:
                    target = float(row[target_index])
                except ValueError:
                    invalid_eligible_targets += 1
                else:
                    if not math.isfinite(target) or target < 0:
                        invalid_eligible_targets += 1
            # Strict field equivalence is stronger than a rounded numerical
            # comparison; header order is checked separately below.
            digest = sha256(json.dumps(row, ensure_ascii=False, separators=(",", ":")).encode()).digest()
            batch.append((representation, row[idx["plant_id"]], key_time, digest, source,
                          int(flag in {"true", "1", "yes"})))
            rows += 1
            if representation == "dataset":
                for name, index in numeric_indices:
                    text = row[index]
                    try:
                        number = float(text)
                    except ValueError:
                        if not text or text.lower() in {"nan", "na", "null"}:
                            numeric_missing[name] += 1
                        else:
                            numeric_invalid[name] += 1
                    else:
                        if number != number:
                            numeric_missing[name] += 1
                        elif abs(number) == float("inf"):
                            # NumericPreprocessor masks nonfinite features;
                            # ineligible raw targets are intentionally retained.
                            numeric_nonfinite[name] += 1
            # Company/year are part of the declared physical partition contract.
            if representation == "partitions":
                if f"company={row[idx['company']]}" not in path.parts or f"year={stamp.year}" not in path.parts:
                    partition_mismatches += 1
            if len(batch) >= batch_rows:
                flush()
        flush()
    return {"path": str(path), "sha256": sha256_file(path), "bytes": path.stat().st_size,
            "rows": rows, "columns": header, "duplicate_keys": duplicates,
            "start": start, "end": end, "energy_rows": dict(energy),
            "eligible_rows": dict(eligible), "numeric_missing": dict(numeric_missing),
            "numeric_invalid": dict(numeric_invalid), "partition_key_mismatches": partition_mismatches,
            "numeric_nonfinite": dict(numeric_nonfinite),
            "invalid_eligible_target_rows": invalid_eligible_targets}


def audit_gold_dataset(manifest_path: Path, *, project_root: Path, batch_rows: int = 5000) -> dict:
    """Compare declared inventory, bytes, rows, schema, unique keys and all fields.

    The legacy manifest has no content checksums. Fresh hashes document this
    inspection; they cannot prove the original build's byte authenticity.
    """
    if batch_rows < 1:
        raise ValueError("batch_rows must be positive")
    root = Path(project_root).resolve()
    manifest_path = _resolve(root, str(manifest_path))
    manifest_sha = sha256_file(manifest_path)
    manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    if manifest.get("contract") != MANIFEST_CONTRACT or manifest.get("schema_version") != 1:
        raise ValueError("Unsupported Gold manifest contract/schema")
    target = manifest["target"]
    features = list(manifest["features"])
    if not features or len(features) != len(set(features)) or target in features:
        raise ValueError("Gold manifest features must be unique and exclude the target")
    dataset = _resolve(root, manifest["dataset"])
    partition_root = _resolve(root, manifest["partitioned_dataset"]["directory"])
    inventory = manifest["partitioned_dataset"]["files"]
    listed = [_resolve(root, entry["path"]) for entry in inventory]
    if not listed or len(listed) != len(set(listed)) or any(not p.is_relative_to(partition_root) for p in listed):
        raise ValueError("Invalid/duplicate Gold partition inventory")
    discovered = DatasetRepository.training_files(partition_root)
    issues = []
    if set(listed) != set(discovered):
        issues.append({"code": "partition_inventory_mismatch",
                       "unlisted": sorted(str(p) for p in set(discovered) - set(listed)),
                       "missing": sorted(str(p) for p in set(listed) - set(discovered))})
    if any(not p.is_file() for p in [dataset, *listed]):
        raise ValueError("A Gold file declared by the manifest is missing")
    before = {p: (p.stat().st_size, p.stat().st_mtime_ns) for p in [dataset, *listed, manifest_path]}
    with TemporaryDirectory(prefix="solar-gold-audit-") as scratch:
        db = sqlite3.connect(str(Path(scratch) / "keys.sqlite"))
        try:
            db.execute("PRAGMA cache_size=-8192")
            db.execute("PRAGMA temp_store=FILE")
            db.execute("PRAGMA journal_mode=OFF")
            db.execute("CREATE TABLE cells (representation TEXT, plant TEXT, timestamp TEXT, digest BLOB, energy TEXT, eligible INTEGER, "
                       "PRIMARY KEY (representation, plant, timestamp)) WITHOUT ROWID")
            combined = _scan_file(dataset, "dataset", db, [*features, target], batch_rows)
            parts = [_scan_file(path, "partitions", db, [*features, target], batch_rows) for path in listed]
            for entry, part in zip(inventory, parts):
                for field in ("rows", "bytes"):
                    if part[field] != entry[field]:
                        issues.append({"code": f"partition_{field}_mismatch", "path": part["path"],
                                       "expected": entry[field], "actual": part[field]})
                if part["columns"] != combined["columns"]:
                    issues.append({"code": "partition_schema_order_mismatch", "path": part["path"]})
                if part["partition_key_mismatches"]:
                    issues.append({"code": "partition_company_year_mismatch", "path": part["path"]})
                if (f"company={entry['company']}" not in Path(part["path"]).parts
                        or f"year={entry['year']}" not in Path(part["path"]).parts):
                    issues.append({"code": "manifest_partition_label_mismatch", "path": part["path"]})
            comparisons = {}
            for a, b in (("dataset", "partitions"), ("partitions", "dataset")):
                comparisons[f"{a}_keys_missing_from_{b}"] = db.execute(
                    "SELECT COUNT(*) FROM cells a LEFT JOIN cells b ON b.representation=? "
                    "AND a.plant=b.plant AND a.timestamp=b.timestamp "
                    "WHERE a.representation=? AND b.plant IS NULL", (b, a)).fetchone()[0]
            comparisons["field_mismatch_rows"] = db.execute(
                "SELECT COUNT(*) FROM cells a JOIN cells b ON a.plant=b.plant AND a.timestamp=b.timestamp "
                "WHERE a.representation='dataset' AND b.representation='partitions' AND a.digest!=b.digest").fetchone()[0]
            plants = db.execute("SELECT COUNT(DISTINCT plant) FROM cells WHERE representation='dataset'").fetchone()[0]
            population = {}
            for energy, plant, eligible in db.execute(
                    "SELECT energy, plant, MAX(eligible) FROM cells WHERE representation='dataset' GROUP BY energy, plant"):
                info = population.setdefault(energy, {"retained_plants": 0, "training_eligible_plants": 0, "fully_excluded_plants": []})
                info["retained_plants"] += 1
                info["training_eligible_plants"] += eligible
                if not eligible:
                    info["fully_excluded_plants"].append(plant)
        finally:
            db.close()
    if any(comparisons.values()):
        issues.append({"code": "export_content_mismatch", **comparisons})
    if combined["duplicate_keys"] or any(p["duplicate_keys"] for p in parts):
        issues.append({"code": "duplicate_plant_timestamp"})
    if sum(p["rows"] for p in parts) != combined["rows"]:
        issues.append({"code": "partition_total_rows_mismatch"})
    for field, actual in (("rows", combined["rows"]), ("plants", plants),
                          ("start", combined["start"]), ("end", combined["end"])):
        if manifest[field] != actual:
            issues.append({"code": f"manifest_{field}_mismatch", "expected": manifest[field], "actual": actual})
    if [c for c in combined["columns"] if c in features] != features:
        issues.append({"code": "manifest_feature_order_mismatch"})
    eligibility = manifest.get("training_eligibility", {})
    if eligibility.get("training_eligible_rows") != sum(combined["eligible_rows"].values()):
        issues.append({"code": "manifest_eligible_rows_mismatch"})
    for source, info in eligibility.get("by_energy_source", {}).items():
        if (info["retained_rows"] != combined["energy_rows"].get(source, 0)
                or info["training_eligible_rows"] != combined["eligible_rows"].get(source, 0)):
            issues.append({"code": "manifest_energy_counts_mismatch", "energy_source": source})
        actual = population.get(source, {"retained_plants": 0, "training_eligible_plants": 0, "fully_excluded_plants": []})
        if any(field in info and info[field] != actual[field] for field in actual):
            issues.append({"code": "manifest_energy_population_mismatch", "energy_source": source})
    if combined["numeric_invalid"]:
        issues.append({"code": "non_numeric_model_values", "columns": combined["numeric_invalid"]})
    if combined["invalid_eligible_target_rows"]:
        issues.append({"code": "invalid_eligible_target", "rows": combined["invalid_eligible_target_rows"]})
    if sha256_file(manifest_path) != manifest_sha or any(
            (p.stat().st_size, p.stat().st_mtime_ns) != state for p, state in before.items()):
        issues.append({"code": "input_changed_during_audit"})
    return {"contract": GOLD_AUDIT_CONTRACT, "status": "FAIL" if issues else "PASS",
            "inputs_read_only": True, "training_performed": False, "model_performance_verified": False,
            "manifest": {"path": str(manifest_path), "sha256": manifest_sha,
                         "build_time_content_hashes_available": False},
            "key_columns": ["plant_id", "timestamp"], "timestamp_policy": "timezone-naive local hourly; no conversion",
            "target": target, "target_unit": "MWh", "features": features,
            "preprocessing_contract": manifest.get("preprocessing_contract"),
            "dataset": combined, "partitions": parts, "plants": plants,
            "population_by_energy_source": population,
            "comparison": comparisons, "issues": issues,
            "scratch_index": "private temporary SQLite; bounded cache; removed after audit"}
