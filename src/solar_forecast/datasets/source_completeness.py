"""Read preserved source files before considering incremental collection."""

from __future__ import annotations

from datetime import date
import json
from pathlib import Path
import re
import warnings

import numpy as np
import pandas as pd

from solar_forecast.collectors.generation_normalizers import (
    ARCHIVE_WIDE_SCHEMAS,
    EwpTrainingNormalizer,
    KOSPO_WIDE_SCHEMA,
    KoenGenerationNormalizer,
    read_csv_with_fallback,
)
from solar_forecast.datasets.archive_standardizer import COMPANY_DIRECTORIES, GenerationSchemaRegistry
from solar_forecast.infrastructure.artifact_store import sha256_file


PROVIDERS = ("koen", "ewp", "kospo")
HOUR = pd.Timedelta(hours=1)
KEY = ["entity", "timestamp"]
SCHEMAS = GenerationSchemaRegistry((*ARCHIVE_WIDE_SCHEMAS, KOSPO_WIDE_SCHEMA))


def _read_source(provider: str, path: Path) -> tuple[str, pd.DataFrame, dict[str, object]]:
    with warnings.catch_warnings():
        if provider == "koen":
            # The monthly export includes an unlabeled, empty trailing field.
            warnings.simplefilter("ignore", pd.errors.ParserWarning)
        raw = read_csv_with_fallback(path, index_col=False)
    raw.columns = raw.columns.str.strip()
    if provider == "ewp" and set(EwpTrainingNormalizer.column_map) <= set(raw.columns):
        layout = "region_hour"
        entity = raw["시도명"].astype("string").str.strip()
        timestamps = pd.to_datetime(raw["발전일자"], errors="coerce", format="mixed")
        values = pd.to_numeric(raw["발전량(MWh)"], errors="coerce")
        hours = pd.DataFrame({"entity": entity, "timestamp": timestamps, "value": values, "unit": "mwh"})
        native_keys = ["시도명", "발전일자"]
        source_unit = "mwh"
    else:
        layout = "plant_meter_hour"
        if provider == "koen":
            day, plant = "일자", "발전구분"
            identity_columns = [plant, "호기"]
            catalog_hours = [f"{hour}시발전량" for hour in range(1, 25)]
            if {"발전일", plant, "호기", *catalog_hours} <= set(raw):
                # Public FileData 15167614 declares kWh for these 24 fields.
                day, columns, source_unit = "발전일", catalog_hours, "kwh"
            else:
                matches = [(column, KoenGenerationNormalizer.hour_pattern.fullmatch(column)) for column in raw]
                matched = {int(match[1]): column for column, match in matches if match}
                declared = {match[2].lower() for _, match in matches if match}
                if set(matched) != set(range(1, 25)) or len(declared) != 1:
                    raise ValueError("KOEN download does not contain one consistent 1–24 hour schema")
                columns = [matched[hour] for hour in range(1, 25)]
                source_unit = declared.pop()
                if source_unit == "mwh" and raw[columns].apply(pd.to_numeric, errors="coerce").max().max() > 1_000:
                    source_unit = "kwh"
        else:
            schema = SCHEMAS.resolve(set(raw), provider)
            day, plant = schema.date_column, schema.plant_column
            identity_columns = [column for column in (plant, schema.unit_column, *schema.id_columns) if column]
            if schema.plant_from_filename:
                name = re.search(r"\[([^]]+)]", path.stem)
                if not name:
                    raise ValueError("Plant name missing from archive filename")
                raw[plant] = name[1].strip()
            columns = list(schema.hour_columns)
            source_unit = schema.source_unit if not schema.source_unit_alternatives else "requires_unit_resolution"
        dates = pd.to_datetime(raw[day].astype("string").str.strip(), errors="coerce", format="mixed").dt.normalize()
        identity = raw[identity_columns].astype("string").apply(lambda column: column.str.strip())
        identity = identity.mask(identity.eq(""))
        if "호기" in identity:
            identity["호기"] = identity["호기"].str.replace(r"\.0$", "", regex=True)
        entities = identity.iloc[:, 0]
        for column in identity_columns[1:]:
            entities = entities.str.cat(identity[column], sep="#")
        values = raw[columns].apply(pd.to_numeric, errors="coerce").to_numpy().ravel()
        offsets = np.tile(np.arange(24), len(raw))
        timestamps = dates.repeat(24).reset_index(drop=True) + pd.to_timedelta(offsets, unit="h")
        hours = pd.DataFrame({"entity": entities.repeat(24).reset_index(drop=True),
                              "timestamp": timestamps, "value": values, "unit": source_unit})
        native_keys = [day, *identity_columns]
    has_entity = hours["entity"].notna() & hours["entity"].ne("")
    on_hour = hours["timestamp"].notna() & hours["timestamp"].eq(hours["timestamp"].dt.floor("h"))
    finite_value = np.isfinite(hours["value"])
    valid = has_entity & on_hour & finite_value
    hours = hours.loc[valid].copy()
    divisor = {"wh": 1_000_000, "kwh": 1_000, "mwh": 1}.get(source_unit)
    if divisor is not None:
        hours["value"] = hours["value"] / divisor
        hours["unit"] = "mwh"
    return layout, hours, {
        "source_rows": len(raw),
        "comparison_source_unit": source_unit,
        "identical_source_rows": int(raw.duplicated().sum()),
        "repeated_source_key_rows": int(raw.duplicated(native_keys).sum()),
        "valid_hour_rows": len(hours),
        "invalid_hour_slots": len(valid) - len(hours),
        "invalid_hour_reasons": {
            "missing_source_identity": int((~has_entity).sum()),
            "invalid_timestamp_or_off_hour": int((has_entity & ~on_hour).sum()),
            "missing_or_nonfinite_value": int((has_entity & on_hour & ~finite_value).sum()),
        },
        "observed_dates": int(hours["timestamp"].dt.normalize().nunique()),
        "first_observed_hour": hours["timestamp"].min().isoformat() if len(hours) else None,
        "last_observed_hour": hours["timestamp"].max().isoformat() if len(hours) else None,
    }


def _hour_gaps(expected: pd.DatetimeIndex, actual: pd.DatetimeIndex) -> list[dict[str, object]]:
    missing = expected.difference(actual)
    if missing.empty:
        return []
    breaks = np.flatnonzero(np.diff(missing.values) != HOUR.to_timedelta64()) + 1
    starts, stops = [0, *breaks], [*breaks, len(missing)]
    return [
        {"first": missing[start].isoformat(), "last": missing[stop - 1].isoformat(), "hours": int(stop - start)}
        for start, stop in zip(starts, stops, strict=True)
    ]


def _timeline(hours: pd.DataFrame, start: date, end: date, *, calendar_months: bool) -> dict[str, object]:
    selected = hours.loc[hours["timestamp"].ge(pd.Timestamp(start)) &
                         hours["timestamp"].lt(pd.Timestamp(end) + pd.Timedelta(days=1))]
    unique_readings = selected.drop_duplicates([*KEY, "unit", "value"])
    repeated = unique_readings.loc[unique_readings.duplicated(KEY, keep=False)]
    comparable = repeated.loc[repeated["unit"].eq("mwh")]
    conflicting = comparable.loc[comparable.duplicated(KEY), KEY].drop_duplicates()
    non_comparable = repeated.loc[repeated["unit"].ne("mwh"), KEY].drop_duplicates()
    candidates = conflicting.groupby("entity").size()
    same_value_repeats = selected["entity"].value_counts() - unique_readings["entity"].value_counts()
    conflict_sample = []
    for entity, timestamp in conflicting.loc[:, KEY].head(12).itertuples(index=False, name=None):
        records = selected.loc[selected["entity"].eq(entity) & selected["timestamp"].eq(timestamp)]
        conflict_sample.append({"entity": entity, "timestamp": timestamp.isoformat(),
                                "source_files": sorted(records["source_file"].unique())})
    coverage = []
    for entity, readings in selected.groupby("entity", sort=True):
        actual = pd.DatetimeIndex(readings["timestamp"].unique()).sort_values()
        expected = pd.date_range(actual.min(), actual.max(), freq="h")
        gaps = _hour_gaps(expected, actual)
        record = {"entity": entity, "first_observed_hour": actual[0].isoformat(),
                  "last_observed_hour": actual[-1].isoformat(), "expected_hours_within_observed_bounds": len(expected),
                  "observed_hours": len(actual), "missing_hours_within_observed_bounds": len(expected) - len(actual),
                  "gaps": gaps, "same_value_repeat_rows": int(same_value_repeats[entity]),
                  "conflicting_value_hours": int(candidates.get(entity, 0))}
        if calendar_months:
            months = []
            for month, timestamps in readings.groupby(readings["timestamp"].dt.to_period("M"))["timestamp"]:
                if month.month not in (5, 9, 10):
                    continue
                month_start = month.start_time
                month_end = month.end_time.floor("h")
                month_axis = pd.date_range(month_start, month_end, freq="h")
                actual_month_axis = pd.DatetimeIndex(timestamps.unique())
                months.append({"month": str(month), "calendar_hours": len(month_axis),
                               "observed_hours": len(actual_month_axis),
                               "gaps": _hour_gaps(month_axis, actual_month_axis)})
            record["observed_calendar_months"] = months
        coverage.append(record)
    return {
        "start_date": start.isoformat(), "end_date": end.isoformat(), "observed_rows": len(selected),
        "unique_hour_keys": len(selected.drop_duplicates(KEY)),
        "same_value_repeat_rows": len(selected) - len(unique_readings),
        "conflicting_value_hour_keys": len(conflicting),
        "conflict_sample": conflict_sample,
        "non_comparable_unit_hour_keys": len(non_comparable),
        "entities": coverage,
    }


def _read_indexes(raw_root: Path, extra_indexes: tuple[Path, ...]) -> list[dict[str, object]]:
    paths = {*raw_root.glob("runs/*/collection_manifest.json"), *extra_indexes}
    indexes = []
    for path in sorted(paths):
        item = {"path": str(path)}
        try:
            with path.open(encoding="utf-8-sig") as source:
                manifest = json.load(source)
            item["sha256"] = sha256_file(path)
            if not isinstance(manifest, dict) or not isinstance(manifest.get("results"), list):
                item["error"] = "Unrecognized collection index schema. Scanned input files remain the coverage evidence."
                if isinstance(manifest, dict):
                    item["top_level_fields"] = sorted(manifest)
                indexes.append(item)
                continue
            results = []
            for result in manifest["results"]:
                if result.get("source") not in PROVIDERS:
                    continue
                evidence = {key: result.get(key) for key in ("source", "status", "rows")}
                if result.get("windows"):
                    fields = ("start_date", "end_date", "status", "source_rows", "detail", "path")
                    evidence["windows"] = [
                        {key: window.get(key) for key in fields} for window in result["windows"]
                    ]
                results.append(evidence)
            item.update({
                "start_date": manifest.get("start_date"),
                "end_date": manifest.get("end_date"),
                "results": results,
                "originals": [
                    {key: artifact.get(key) for key in ("path", "sha256")}
                    for artifact in manifest.get("file_artifacts", [])
                    if artifact.get("role") == "provider_original_bronze"
                ],
            })
        except (OSError, KeyError, TypeError, AttributeError, ValueError) as exc:
            item["error"] = str(exc)
        indexes.append(item)
    return indexes


def audit_source_completeness(
    archive_root: Path, raw_root: Path, as_of: date, *, collection_indexes: tuple[Path, ...] = (),
) -> dict[str, object]:
    """Compare provider evidence without interpreting an audit gap as missing published data."""
    indexes = _read_indexes(raw_root, collection_indexes)
    indexed_hashes = {entry["sha256"] for index in indexes for entry in index.get("originals", [])}
    providers = {}
    for provider in PROVIDERS:
        archive = archive_root / COMPANY_DIRECTORIES[provider]
        bronze = raw_root / provider
        files, layouts, seen = [], {}, {}
        for source_root in (archive, bronze):
            for path in sorted(source_root.rglob("*.csv")):
                if "normalized" in path.relative_to(source_root).parts:
                    continue
                checksum = sha256_file(path)
                item = {"path": str(path), "sha256": checksum, "bytes": path.stat().st_size,
                        "indexed_by_collection_manifest": checksum in indexed_hashes}
                files.append(item)
                filename_plant = re.search(r"\[([^]]+)]", path.stem) if provider == "kospo" else None
                replay_key = (checksum, filename_plant[1] if filename_plant else None)
                if replay_key in seen:
                    item["duplicate_of"] = seen[replay_key]
                    continue
                seen[replay_key] = str(path)
                try:
                    layout, hours, counts = _read_source(provider, path)
                except (KeyError, TypeError, ValueError, OSError) as exc:
                    item["read_error"] = str(exc)
                    continue
                item.update(counts, layout=layout, status="read" if len(hours) else "no_value_rows")
                since_2024 = hours.loc[hours["timestamp"].ge("2024-01-01")].assign(source_file=str(path))
                layouts.setdefault(layout, []).append(since_2024)
        coverage_by_layout = {}
        for layout, batches in layouts.items():
            hours = pd.concat(batches, ignore_index=True)
            coverage_by_layout[layout] = {
                str(year): _timeline(hours, date(year, 1, 1), min(as_of, date(year, 12, 31)),
                                     calendar_months=provider == "koen" and year == 2024)
                for year in (2024, 2025, 2026) if date(year, 1, 1) <= as_of
            }
        providers[provider] = {"files": files, "coverage_by_layout": coverage_by_layout}
    scanned_hashes = {item["sha256"] for provider in providers.values() for item in provider["files"]}
    for index in indexes:
        for entry in index.get("originals", []):
            entry["hash_present_in_scanned_inputs"] = entry["sha256"] in scanned_hashes
    return {
        "contract": "solar-source-completeness.v1", "as_of_kst": as_of.isoformat(),
        "timestamp_basis": "source hourly labels on the repository's KST axis",
        "gap_policy": "Observed bounds and calendar comparisons are diagnostic. An absent hour does not establish publication or asset operation.",
        "revision_policy": "Conflicting values are review candidates. Authority and publication order are required before applying a correction.",
        "archive_root": str(archive_root), "archive_root_exists": archive_root.is_dir(),
        "bronze_root": str(raw_root), "bronze_root_exists": raw_root.is_dir(),
        "collection_indexes": indexes, "providers": providers,
    }
