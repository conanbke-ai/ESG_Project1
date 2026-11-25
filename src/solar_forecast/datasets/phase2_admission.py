"""Audit retained 2026 observations and version their Gold admission decision."""
from __future__ import annotations

import csv
from dataclasses import dataclass
import gzip
from itertools import zip_longest
import json
from pathlib import Path

import numpy as np
import pandas as pd

from solar_forecast.collectors.generation_normalizers import (
    GENERATION_COLUMNS,
    KOSPO_WIDE_SCHEMA,
    DailyWideGenerationNormalizer,
    EwpTrainingNormalizer,
    KoenGenerationNormalizer,
    read_csv_with_fallback,
)
from solar_forecast.collectors.plant_identity import resolve_generation_identity
from solar_forecast.collectors.plant_metadata import PlantMetadataCatalog
from solar_forecast.datasets.archive_standardizer import HistoricalGenerationStandardizationService
from solar_forecast.datasets.gold_audit import GOLD_AUDIT_CONTRACT, audit_gold_dataset
from solar_forecast.datasets.plant_registry import KmaStationCatalog, ReviewedStationMappingCatalog
from solar_forecast.datasets.repository import DatasetRepository
from solar_forecast.features.asos_features import KmaAsosNormalizer
from solar_forecast.features.future_weather import (
    ALLOWED_COORDINATE_SOURCES,
    FUTURE_WEATHER_COORDINATE_COLUMNS,
    FUTURE_WEATHER_KEYS,
    future_weather_feature_columns,
)
from solar_forecast.infrastructure.artifact_store import sha256_file
from solar_forecast.infrastructure.project_paths import (
    GENERATION_ARCHIVE_ROOT,
    GOLD_DATASET_PATH,
    GOLD_PARTITIONS_ROOT,
    WEATHER_ARCHIVE_ROOT,
)
from solar_forecast.quality.generation_quality import GenerationQualityPolicy, WEATHER_RANGES


CONTRACT = "solar-gold-candidate-admission.v1"
DATA_VERSION = "gold-2026-phase2-v1"
BASELINE_ROWS = 718_531
BASELINE_COLUMNS = 96
BASELINE_PARTITIONS = 27
HORIZONS = (1, 24, 72)
BASELINE_MANIFEST = GOLD_DATASET_PATH.with_name("model_ready_manifest.json")
BASELINE_VERIFICATION = Path("artifacts/verification/architecture_convergence/20261125/actual_gold_candidate.json")
WEATHER_STAGING = Path("artifacts/weather_full_versions/20261104_weather_full_staging_v1/manifest.json")
RETAINED_ARCHIVE = GENERATION_ARCHIVE_ROOT / "한국남부발전" / "최신공공데이터"
RETAINED_GENERATION = (
    ("[용수리]태양광발전실적_20260228.csv", "15095404"),
    ("[신풍리]태양광발전실적_20260228.csv", "15109998"),
)
AVAILABILITY_COLUMNS = (
    "plant_id", "timestamp", "generation_available_at", "weather_available_at",
    "generation_evidence", "weather_evidence",
)


@dataclass(frozen=True)
class GenerationSource:
    company: str
    path: Path
    catalog_url: str | None = None
    catalog_registration_date: str | None = None


def retained_generation_sources() -> tuple[GenerationSource, ...]:
    return tuple(
        GenerationSource("kospo", RETAINED_ARCHIVE / filename,
                         f"https://www.data.go.kr/data/{dataset}/fileData.do", "2026-04-09")
        for filename, dataset in RETAINED_GENERATION
    )


def _indexed_generation(root: Path, audit_path: Path) -> tuple[GenerationSource, ...]:
    report = json.loads((root / audit_path).read_text(encoding="utf-8-sig"))
    if report.get("contract") != "solar-source-completeness.v1":
        raise ValueError("Unsupported source completeness report")
    sources = []
    for company, result in report["providers"].items():
        for record in result["files"]:
            first = record.get("first_observed_hour")
            last = record.get("last_observed_hour")
            if not first or not last or first[:4] > "2026" or last[:4] < "2026":
                continue
            path = Path(record["path"])
            artifact = _artifact(root, path)
            if not artifact["sha256"] or record["sha256"] != artifact["sha256"]:
                raise ValueError(f"Source differs from the reused collection evidence: {path}")
            sources.append(GenerationSource(company, path))
    return tuple(sources)


def _artifact(root: Path, path: Path) -> dict:
    path = root / path
    label = _artifact_label(root, path)
    if not path.is_file():
        return {"path": label, "status": "missing", "bytes": None, "sha256": None}
    return {"path": label, "status": "present", "bytes": path.stat().st_size,
            "sha256": sha256_file(path)}


def _reuse_verified_gold(root: Path, manifest: Path, verification: Path) -> dict | None:
    if not verification.is_file():
        return None
    try:
        result = json.loads(verification.read_text(encoding="utf-8-sig"))
        if (
            result.get("contract") != GOLD_AUDIT_CONTRACT
            or result.get("status") != "PASS"
            or result.get("issues") != []
            or result["manifest"]["sha256"] != sha256_file(manifest)
        ):
            return None
        contract = json.loads(manifest.read_text(encoding="utf-8-sig"))
        combined = (root / contract["dataset"]).resolve()
        partition_root = (root / contract["partitioned_dataset"]["directory"]).resolve()
        listed = {(root / item["path"]).resolve() for item in contract["partitioned_dataset"]["files"]}
        checked_partitions = {(root / item["path"]).resolve() for item in result["partitions"]}
        if (
            combined != (root / result["dataset"]["path"]).resolve()
            or listed != checked_partitions
            or len(checked_partitions) != len(result["partitions"])
        ):
            return None
        if not partition_root.is_relative_to(root) or set(DatasetRepository.training_files(partition_root)) != listed:
            return None
        for item in [result["dataset"], *result["partitions"]]:
            path = (root / item["path"]).resolve()
            if (
                not path.is_relative_to(root) or not path.is_file()
                or item["bytes"] != path.stat().st_size or item["sha256"] != sha256_file(path)
            ):
                return None
        return result
    except (AttributeError, KeyError, OSError, TypeError, ValueError):
        return None


def _baseline(root: Path, manifest: Path, verification: Path) -> dict:
    path = root / manifest
    report = {
        "documented_rows": BASELINE_ROWS,
        "expected_columns": BASELINE_COLUMNS,
        "expected_partitions": BASELINE_PARTITIONS,
        "published_reference": _artifact(root, Path("docs/DATA_COLLECTION_AUDIT.md")),
        "manifest": _artifact(root, manifest),
        "stored_verification": _artifact(root, verification),
        "combined_file": _artifact(root, GOLD_DATASET_PATH),
        "partition_files_present": sum(1 for p in (root / GOLD_PARTITIONS_ROOT).rglob("*.csv.gz") if p.is_file()),
        "verified_rows": None,
    }
    if not path.is_file():
        return {**report, "status": "physical_manifest_missing"}
    try:
        stored = _reuse_verified_gold(root, path, root / verification)
        checked = stored or audit_gold_dataset(path, project_root=root)
    except (KeyError, OSError, ValueError) as error:
        return {**report, "status": "physical_audit_failed", "error": str(error)}
    rows, end = checked["dataset"]["rows"], checked["dataset"]["end"]
    columns, partitions = len(checked["dataset"]["columns"]), len(checked["partitions"])
    valid = (
        checked["status"] == "PASS" and rows == BASELINE_ROWS and columns == BASELINE_COLUMNS
        and partitions == BASELINE_PARTITIONS and end is not None and end < "2026-01-01"
    )
    return {**report, "status": "verified" if valid else "baseline_mismatch",
            "verified_rows": rows, "end": end, "issues": checked["issues"],
            "combined_sha256": checked["dataset"]["sha256"],
            "verified_dataset": _artifact_label(root, root / checked["dataset"]["path"]),
            "verified_columns": columns, "verified_partitions": partitions,
            "verification_method": "stored_audit_with_current_content_hashes" if stored else "read_only_gold_audit"}


def _artifact_label(root: Path, path: Path) -> str:
    return path.relative_to(root).as_posix() if path.is_relative_to(root) else str(path)


def _text_present(values: pd.Series) -> pd.Series:
    return ~values.fillna("").astype(str).str.strip().str.casefold().isin({"", "none", "nan", "null"})


def _kst_local(values: pd.Series) -> pd.Series:
    return pd.to_datetime(values, format="mixed", errors="coerce")


def _actual_time(values: pd.Series) -> pd.Series:
    explicit_zone = values.astype("string").str.contains(r"(?:Z|[+-]\d{2}:\d{2})$", regex=True, na=False)
    utc = pd.to_datetime(values.where(explicit_zone), format="mixed", errors="coerce", utc=True)
    return utc.dt.tz_convert("Asia/Seoul").dt.tz_localize(None)


def _columns(path: Path, required: tuple[str, ...]) -> pd.DataFrame:
    frame = read_csv_with_fallback(path, usecols=lambda col: col in required)
    if missing := set(required) - set(frame.columns):
        raise ValueError(f"Required columns missing from {path.name}: {sorted(missing)}")
    return frame


def _review_evidence(root: Path, review: dict) -> dict[str, dict]:
    """Resolve retained evidence IDs against their source bytes, without remote fetching."""
    verified = {}
    for key, entry in review.get("evidence_sources", {}).items():
        path = (root / entry["path"]).resolve()
        if not path.is_relative_to(root):
            raise ValueError(f"Review evidence must be inside the data checkout: {key}")
        artifact = _artifact(root, path)
        if artifact["sha256"] and artifact["sha256"] == entry.get("sha256"):
            verified[key] = artifact
    return verified


def _daily_totals(raw: pd.DataFrame) -> dict | None:
    hours = [str(hour) for hour in range(1, 25)]
    if not {"년월일", "총량(kw)", *hours}.issubset(raw.columns):
        return None
    dates = _kst_local(raw["년월일"].astype(str))
    same_year = raw.loc[dates.dt.year.eq(2026)]
    hour_sum = same_year[hours].apply(pd.to_numeric, errors="coerce").sum(axis=1, min_count=24)
    total = pd.to_numeric(same_year["총량(kw)"], errors="coerce")
    delta = hour_sum.sub(total).abs()
    return {"days": len(same_year), "comparable_days": int(delta.notna().sum()),
            "mismatch_days": int(delta.gt(0.001).sum()),
            "maximum_absolute_difference_source_scale": float(delta.max()) if delta.notna().any() else None}


def _read_generation(
    root: Path, source: GenerationSource, standardizer: HistoricalGenerationStandardizationService
) -> tuple[dict, pd.DataFrame]:
    artifact = _artifact(root, source.path)
    evidence = {**artifact, "company": source.company, "catalog_url": source.catalog_url,
                "catalog_registration_date": source.catalog_registration_date}
    if artifact["status"] == "missing":
        return evidence, pd.DataFrame()
    path = root / source.path
    try:
        raw = read_csv_with_fallback(path, index_col=False)
        raw.columns = raw.columns.str.strip()
        if set(GENERATION_COLUMNS).issubset(raw):
            frame = raw[GENERATION_COLUMNS].copy()
        elif source.company == "kospo" and set(KOSPO_WIDE_SCHEMA.hour_columns).issubset(raw):
            frame = DailyWideGenerationNormalizer(KOSPO_WIDE_SCHEMA).transform(raw, source_file=path.name)
        elif source.company == "koen" and {"발전일", *(f"{hour}시발전량" for hour in range(1, 25))}.issubset(raw):
            # FileData 15167614 declares kWh for this schema's unitless hourly headings.
            renamed = raw.rename(columns={"발전일": "일자", **{
                f"{hour}시발전량": f"{hour}시 발전량(kWh)" for hour in range(1, 25)
            }})
            frame = KoenGenerationNormalizer().transform(renamed, source_file=path.name)
        elif source.company == "ewp" and set(EwpTrainingNormalizer.column_map).issubset(raw):
            return ({**evidence, "status": "region_aggregate_only", "raw_rows": len(raw),
                     "raw_columns": list(raw), "reason": "Region-hour rows lack plant, generator and meter keys"},
                    pd.DataFrame())
        else:
            frame = standardizer.normalize_source(source.company, path)
        if not frame["company"].eq(source.company).all():
            raise ValueError("Normalized company differs from the declared source company")
        frame = resolve_generation_identity(frame)
        units = frame.attrs.get("generation_unit_resolution")
        frame = standardizer.metadata.enrich(frame)
        stamps = _kst_local(frame["timestamp"])
        if not pd.api.types.is_datetime64_any_dtype(stamps.dtype) or stamps.dt.tz is not None:
            raise ValueError("Generation timestamps must follow the naive KST Gold key convention")
        if stamps.isna().any() or stamps.ne(stamps.dt.floor("h")).any():
            raise ValueError("Generation timestamps must be valid hourly keys")
        frame["timestamp"] = stamps
        frame["generation_mwh"] = pd.to_numeric(frame["generation_mwh"], errors="coerce")
        year = frame.loc[stamps.dt.year.eq(2026)].copy()
        year["_source"] = artifact["path"]
        year["_unit_resolved"] = bool(units and units.get("resolved_source_unit"))
        year["_interval_explicit"] = "0-1시 발전량(kWh)" in raw
        year["plant"] = year["plant"].map(lambda label: standardizer.metadata.canonical_plant(source.company, label))
        year["physical_plant_id"] = source.company + ":" + year["plant"]
        year.attrs = {}
        return ({
            **evidence, "status": "parsed", "raw_rows": len(raw),
            "normalized_all_rows": len(frame), "normalized_2026_rows": len(year),
            "raw_columns": list(raw), "unit_resolution": units, "daily_total_check": _daily_totals(raw),
            "meter_category_values": (
                sorted(raw["계량구분"].dropna().astype(str).unique()) if "계량구분" in raw else []
            ),
        }, year)
    except (OSError, ValueError, TypeError) as error:
        return {**evidence, "status": "invalid", "reason": str(error)}, pd.DataFrame()


def _forecast_columns(profile: str) -> tuple[str, ...]:
    return (*FUTURE_WEATHER_KEYS, *future_weather_feature_columns(profile),
            *FUTURE_WEATHER_COORDINATE_COLUMNS, "forecast_source", "forecast_model")


def _read_forecast(path: Path, profile: str) -> pd.DataFrame:
    columns = {*_forecast_columns(profile), "forecast_issued_at", "forecast_received_at"}
    return read_csv_with_fallback(path, usecols=lambda col: col in columns)


def _forecast(path: Path | None, expected: pd.DatetimeIndex, plant_id: str, horizon: int,
              feature_profile: str, minimum_coverage: float, *, archive: pd.DataFrame | None = None) -> dict:
    summary = {"expected_targets": len(expected), "archive_rows": 0, "aligned_targets": 0,
               "verified_timing_targets": 0, "coverage": 0.0, "minimum_coverage": minimum_coverage}
    if path is None:
        return {**summary, "status": "archive_not_configured"}
    if not path.is_file():
        return {**summary, "status": "archive_missing"}
    features = future_weather_feature_columns(feature_profile)
    actual = ("forecast_issued_at", "forecast_received_at")
    columns = _forecast_columns(feature_profile)
    frame = archive if archive is not None else _read_forecast(path, feature_profile)
    if missing := set(columns) - set(frame.columns):
        return {**summary, "status": "missing_forecast_columns", "missing_columns": sorted(missing)}
    selected = frame.loc[frame["plant_id"].eq(plant_id)
                         & pd.to_numeric(frame["horizon_hours"], errors="coerce").eq(horizon)].copy()
    selected["timestamp"] = _kst_local(selected["timestamp"])
    selected["forecast_origin"] = _kst_local(selected["forecast_origin"])
    selected = selected.loc[selected["timestamp"].isin(expected)]
    summary["archive_rows"] = len(selected)
    duplicate = selected["timestamp"].duplicated(keep=False)
    nominal = selected["timestamp"] - pd.Timedelta(hours=horizon)
    payload = selected[list(features)].apply(pd.to_numeric, errors="coerce")
    coordinates = {
        "weather_query_latitude": (32.0, 39.5), "weather_query_longitude": (124.0, 132.5),
        "grid_latitude": (-90.0, 90.0), "grid_longitude": (-180.0, 180.0),
    }
    spatial = selected["coordinate_source"].isin(ALLOWED_COORDINATE_SOURCES) & selected["cell_selection"].eq("nearest")
    for column, (lower, upper) in coordinates.items():
        spatial &= pd.to_numeric(selected[column], errors="coerce").between(lower, upper)
    aligned = (
        ~duplicate & selected["forecast_origin"].eq(nominal) & np.isfinite(payload).any(axis=1)
        & spatial & _text_present(selected["forecast_model"]) & _text_present(selected["forecast_source"])
    )
    summary["aligned_targets"] = int(aligned.sum())
    summary["duplicate_rows"] = int(duplicate.sum())
    summary["aligned_coverage"] = summary["aligned_targets"] / len(expected)
    if not set(actual).issubset(selected):
        return {**summary, "status": "actual_issue_or_receipt_missing",
                "missing_columns": sorted(set(actual) - set(selected.columns))}
    issued, received = (_actual_time(selected[column]) for column in actual)
    safe = aligned & issued.le(received) & received.le(nominal)
    summary["verified_timing_targets"] = int(safe.sum())
    summary["coverage"] = summary["verified_timing_targets"] / len(expected)
    return {**summary, "status": "verified" if summary["coverage"] >= minimum_coverage else "insufficient_as_of_coverage"}


def _observation_coverage(
    availability: pd.DataFrame, plant_id: str, expected: pd.DatetimeIndex, evidence_ids: set[str],
) -> dict:
    counts = {"expected_targets": len(expected), "generation_timestamp_rows": 0, "weather_timestamp_rows": 0}
    if availability.empty:
        return {**counts, "status": "actual_observation_availability_missing"}
    rows = availability.loc[availability["plant_id"].eq(plant_id)].copy()
    rows["timestamp"] = _kst_local(rows["timestamp"])
    rows = rows.loc[rows["timestamp"].isin(expected)]
    unique = ~rows["timestamp"].duplicated(keep=False)
    for kind in ("generation", "weather"):
        actual = _actual_time(rows[f"{kind}_available_at"])
        earliest = rows["timestamp"] + pd.Timedelta(hours=1) if kind == "generation" else rows["timestamp"]
        valid = unique & actual.ge(earliest) & rows[f"{kind}_evidence"].isin(evidence_ids)
        counts[f"{kind}_timestamp_rows"] = int(valid.sum())
    verified = all(counts[f"{kind}_timestamp_rows"] == len(expected) for kind in ("generation", "weather"))
    return {**counts, "status": "verified" if verified else "incomplete_actual_observation_availability"}


def _plant_report(
    group: pd.DataFrame, plant_id: str, stations: KmaStationCatalog,
    mappings: ReviewedStationMappingCatalog, observed: pd.DataFrame, availability: pd.DataFrame,
    forecasts: dict[int, tuple[Path | None, str]], forecast_coverage: float,
    archives: dict[int, pd.DataFrame], review: dict, evidence_ids: set[str],
) -> tuple[dict, pd.DataFrame]:
    first = group.iloc[0]
    start, end = group["timestamp"].min(), group["timestamp"].max()
    expected = pd.date_range(start, end, freq="h")
    grouped = group.groupby(["plant_id", "timestamp"], observed=True, dropna=False)
    conflicting = int(grouped["generation_mwh"].nunique().gt(1).sum())
    source_series = group.drop_duplicates(["plant_id", "timestamp"])
    coverage_per_series = source_series.groupby("plant_id", observed=True)["timestamp"].nunique()
    series_ids = sorted(str(value) for value in group["plant_id"].unique())
    generator_ids = sorted(group.loc[_text_present(group["unit"]), "unit"].astype(str).unique())
    capacity = pd.to_numeric(group["capacity_mw"], errors="coerce")
    capacities = sorted(capacity.loc[np.isfinite(capacity) & capacity.gt(0)].unique().tolist())
    address = first["address"] if _text_present(group["address"]).all() else None
    mapping = mappings.get(first["company"], first["plant"])
    lat, lon = first["latitude"], first["longitude"]
    match = stations.resolve(
        address=address, latitude=lat, longitude=lon,
        reviewed_station_id=mapping.station_id if mapping else None,
        generation_start=start, generation_end=end,
    )
    measured = (
        source_series.groupby("timestamp", as_index=False, sort=True)["generation_mwh"].sum(min_count=1)
        .assign(plant_id=plant_id, energy_source=first["energy_source"],
                capacity_mw=capacities[0] if len(capacities) == 1 else np.nan)
    )
    weather = observed.loc[observed["station_id"].eq(match.station_id)]
    measured = measured.merge(weather, on="timestamp", how="left", validate="one_to_one")
    observed_keys = int(measured["weather_station_hour_present"].eq(True).sum())
    quality = GenerationQualityPolicy().apply(measured)
    physical = (
        np.isfinite(quality["generation_mwh"]) & quality["quality_train_eligible"]
        & ~quality["quality_capacity_exceeded"]
    )
    observation = _observation_coverage(availability, plant_id, expected, evidence_ids)
    forecast = {str(h): _forecast(path, expected, plant_id, h, profile, forecast_coverage, archive=archives.get(h))
                for h, (path, profile) in forecasts.items()}
    meter_points = review.get("meter_points", {})
    point_ids = [str(meter_points.get(series, {}).get("point_id") or "").strip() for series in series_ids]
    meter_confirmed = all(
        meter_points.get(series, {}).get("evidence") in evidence_ids
        for series in series_ids
    ) and all(point_ids) and len(set(point_ids)) == len(series_ids)
    interval_evidence = review.get("kst_interval_start_evidence")
    unit_evidence = review.get("generation_unit_evidence")
    checks = {
        "plant_identity": bool(address or (pd.notna(lat) and pd.notna(lon))),
        "generator_identity": bool(_text_present(group["unit"]).all()),
        "meter_point_identity": meter_confirmed,
        "generation_mwh": bool(
            (group["_unit_resolved"].all() or unit_evidence in evidence_ids)
            and np.isfinite(group["generation_mwh"]).all()
        ),
        "capacity_mw": bool(len(capacities) == 1 and capacity.gt(0).all() and np.isfinite(capacity).all()),
        "kst_interval_start": bool(group["_interval_explicit"].all() or interval_evidence in evidence_ids),
        "hourly_generation_coverage": bool(conflicting == 0 and coverage_per_series.eq(len(expected)).all()),
        "actual_observation_availability": observation["status"] == "verified",
        "location_station_mapping": bool(match.station_id is not None and not match.review_required),
        "matched_asos_station_hours": observed_keys == len(expected),
        "generation_quality": bool(physical.all()),
        "forecast_horizons_and_timing": all(result["status"] == "verified" for result in forecast.values()),
    }
    finite = quality["generation_mwh"].where(np.isfinite(quality["generation_mwh"]))
    report = {
        "plant_id": plant_id, "company": first["company"], "plant": first["plant"],
        "generation_sources": sorted(group["_source"].unique()),
        "source_series_ids": series_ids, "generator_labels": generator_ids, "meter_points": meter_points,
        "interval_start_evidence": interval_evidence,
        "provisional_start_kst": start.isoformat(), "provisional_end_kst_inclusive": end.isoformat(),
        "normalized_source_rows": len(group), "duplicate_series_hour_rows": len(group) - len(source_series),
        "conflicting_series_hours": conflicting, "expected_hours_per_series": len(expected),
        "observed_hours_by_series": {str(key): int(value) for key, value in coverage_per_series.items()},
        "distinct_plant_hours": len(measured), "as_of_observations": observation,
        "address": address, "capacity_mw_values": capacities,
        "min_generation_mwh": float(finite.min()) if finite.notna().any() else None,
        "max_generation_mwh": float(finite.max()) if finite.notna().any() else None,
        "station": {"id": match.station_id, "name": match.station_name, "method": match.method,
                    "evidence_url": mapping.evidence_url if mapping else None,
                    "rationale": mapping.rationale if mapping else None},
        "matched_asos_station_hours": observed_keys,
        "generation_quality_eligible_rows": int(quality["quality_train_eligible"].sum()),
        "physical_quality_eligible_rows": int(physical.sum()), "forecasts": forecast,
        "checks": checks, "blocked_by": [name for name, value in checks.items() if not value],
    }
    if unit_evidence:
        report["generation_unit_evidence"] = unit_evidence
    return report, measured.assign(station_id=match.station_id)


def _read_gold_rows(path: Path):
    opener = gzip.open if path.name.endswith(".gz") else open
    with opener(path, mode="rt", encoding="utf-8-sig", newline="") as stream:
        yield from csv.reader(stream, strict=True)


def _materialized_candidate(
    root: Path, candidate_manifest: Path | None, baseline: dict,
    expected: pd.DataFrame, reports: list[dict],
) -> dict:
    result = {
        "status": "blocked", "admitted_2026_rows": 0, "dataset_path": None,
        "dataset_sha256": None, "eligible_periods": [],
        "plants_passing_input_checks": [p["plant_id"] for p in reports if not p["blocked_by"]],
        "blockers": [
            *([] if baseline["status"] == "verified" else ["baseline_gold_not_physically_verified"]),
            *([] if not expected.empty else ["no_plant_passes_input_checks"]),
        ],
    }
    if candidate_manifest is None:
        result["blockers"].append("versioned_gold_dataset_not_materialized")
        return result
    result["versioned_gold_manifest"] = _artifact(root, candidate_manifest)
    if result["blockers"]:
        return result
    candidate_dir = (root / candidate_manifest).resolve().parent
    original = (root / baseline["verified_dataset"]).resolve()
    if candidate_dir.name != DATA_VERSION or original.is_relative_to(candidate_dir):
        raise ValueError(f"Use a new {DATA_VERSION}/ directory for the candidate Gold manifest")
    audit = audit_gold_dataset(candidate_manifest, project_root=root)
    dataset = Path(audit["dataset"]["path"])
    if audit["status"] != "PASS":
        raise ValueError(f"Candidate Gold consistency audit failed: {audit['issues']}")
    if not all(Path(item["path"]).is_relative_to(candidate_dir)
               for item in [audit["dataset"], *audit["partitions"]]):
        raise ValueError("Every candidate Gold file must be in the versioned directory")
    if audit["dataset"]["rows"] != BASELINE_ROWS + len(expected):
        raise ValueError("Candidate Gold count differs from the baseline plus eligible extension")
    old_rows, candidate_rows = _read_gold_rows(original), _read_gold_rows(dataset)
    expected_header = next(old_rows)
    if next(candidate_rows) != expected_header:
        raise ValueError("Candidate Gold changes the verified baseline schema/order")
    time_index = expected_header.index("timestamp")
    old_candidate_rows = (row for row in candidate_rows if row[time_index] < "2026-01-01")
    if any(old != new for old, new in zip_longest(old_rows, old_candidate_rows)):
        raise ValueError("Candidate Gold changes the verified baseline rows")
    keys = ["plant_id", "timestamp"]
    values = ["generation_mwh", *[name for name in ["capacity_mw", *WEATHER_RANGES] if name in expected]]
    columns = [*keys, *values, "station_id", "quality_train_eligible"]
    extension = []
    for chunk in pd.read_csv(dataset, usecols=columns, chunksize=100_000):
        chunk["timestamp"] = _kst_local(chunk["timestamp"])
        extension.append(chunk.loc[chunk["timestamp"].dt.year.eq(2026)])
    actual = pd.concat(extension, ignore_index=True)
    joined = expected.merge(actual, on=keys, how="outer", suffixes=("_source", "_candidate"),
                            validate="one_to_one", indicator=True)
    if (
        not joined["_merge"].eq("both").all()
        or not joined["quality_train_eligible"].astype(str).str.casefold().isin({"true", "1", "yes"}).all()
        or not joined["station_id_source"].eq(joined["station_id_candidate"]).all()
        or any(not np.allclose(joined[f"{name}_source"], joined[f"{name}_candidate"],
                              rtol=1e-12, atol=1e-12, equal_nan=True) for name in values)
    ):
        raise ValueError("Candidate Gold 2026 rows differ from audited source/station/quality evidence")
    if sha256_file(original) != baseline["combined_sha256"] or sha256_file(dataset) != audit["dataset"]["sha256"]:
        raise ValueError("A Gold file changed during candidate verification")
    periods = [
        {"plant_id": str(plant), "start": rows["timestamp"].min().isoformat(),
         "end": rows["timestamp"].max().isoformat(), "rows": len(rows)}
        for plant, rows in actual.groupby("plant_id", sort=True)
    ]
    return {**result, "status": "admitted", "admitted_2026_rows": len(actual),
            "dataset_path": _artifact_label(root, dataset), "dataset_sha256": audit["dataset"]["sha256"],
            "eligible_periods": periods, "period_end_inclusive": True}


def audit_2026_gold_candidate(
    root: Path, *, generation_sources: tuple[GenerationSource, ...] = (),
    availability_path: Path | None = None, review_path: Path | None = None,
    forecast_paths: dict[int, Path] | None = None,
    weather_csv: Path | None = None, weather_staging_manifest: Path = WEATHER_STAGING,
    baseline_manifest: Path = BASELINE_MANIFEST, verification: Path = BASELINE_VERIFICATION,
    candidate_manifest: Path | None = None,
    source_audit: Path | None = None,
) -> dict:
    """Read evidence and emit an admission decision, without rerunning Gold or training."""
    root = root.resolve()
    generation_root = root / GENERATION_ARCHIVE_ROOT
    weather_root = root / WEATHER_ARCHIVE_ROOT
    station_path = weather_root / "META_관측지점정보.csv"
    weather_path = root / weather_csv if weather_csv else weather_root / "OBS_ASOS_TIM_2026.csv"
    mappings_path = root / "config/reviewed_weather_mappings.json"
    experiment_path = root / "config/experiments/optimized.json"
    settings = json.loads(experiment_path.read_text(encoding="utf-8"))["future_weather"]
    defaults = settings["by_horizon"]
    forecasts = {}
    for horizon in HORIZONS:
        config = defaults.get(str(horizon), {})
        source = (forecast_paths or {}).get(horizon, config.get("source"))
        forecasts[horizon] = (root / source if source else None,
                              config.get("feature_profiles", ["aligned_core"])[0])
    metadata = PlantMetadataCatalog.from_directory(generation_root / "location")
    standardizer = HistoricalGenerationStandardizationService(generation_root, root, metadata=metadata)
    indexed = _indexed_generation(root, source_audit) if source_audit else ()
    sources = {
        (source.company, (root / source.path).resolve()): source
        for source in (*indexed, *generation_sources, *retained_generation_sources())
    }
    generation = [_read_generation(root, source, standardizer) for source in sources.values()]
    nonempty = [frame for _, frame in generation if not frame.empty]
    plants = pd.concat(nonempty, ignore_index=True) if nonempty else pd.DataFrame()
    availability = (
        _columns(root / availability_path, AVAILABILITY_COLUMNS)
        if availability_path and (root / availability_path).is_file() else pd.DataFrame()
    )
    review = json.loads((root / review_path).read_text(encoding="utf-8")) if review_path else {}
    review_evidence = _review_evidence(root, review)
    stations = KmaStationCatalog.from_metadata(station_path)
    mappings = ReviewedStationMappingCatalog.from_json(mappings_path)
    observations = (
        KmaAsosNormalizer(station_path).read([weather_path])
        if weather_path.is_file() else pd.DataFrame(columns=["station_id", "timestamp", "weather_station_hour_present"])
    )
    archives = {h: _read_forecast(path, profile) for h, (path, profile) in forecasts.items()
                if path and path.is_file() and not plants.empty}
    audited = [
        _plant_report(group, str(plant), stations, mappings, observations, availability,
                      forecasts, settings["minimum_coverage"], archives,
                      review.get("plants", {}).get(str(plant), {}), set(review_evidence))
        for plant, group in plants.groupby("physical_plant_id", sort=True)
    ] if not plants.empty else []
    plant_reports = [report for report, _ in audited]
    passing = [rows[["plant_id", "timestamp", "generation_mwh", "station_id", "capacity_mw", *WEATHER_RANGES]]
               for report, rows in audited if not report["blocked_by"]]
    baseline = _baseline(root, baseline_manifest, verification)
    return {
        "contract": CONTRACT, "data_version": DATA_VERSION, "year": 2026,
        "time_zone": "Asia/Seoul", "horizons_hours": list(HORIZONS),
        "baseline_gold": baseline,
        "inputs": {
            "generation_sources": [source for source, _ in generation],
            **({"source_completeness": _artifact(root, source_audit)} if source_audit else {}),
            "official_plant_metadata": [_artifact(root, path) for path in sorted((generation_root / "location").glob("*.csv"))],
            "station_metadata": _artifact(root, station_path),
            "station_mapping": _artifact(root, mappings_path),
            "asos_2026": _artifact(root, weather_path),
            "weather_staging_manifest": _artifact(root, weather_staging_manifest),
            "experiment_config": _artifact(root, experiment_path),
            "observation_availability": _artifact(root, availability_path) if availability_path else None,
            "plant_review": _artifact(root, review_path) if review_path else None,
            **({"verified_review_sources": review_evidence} if review else {}),
            "forecast_archives": {str(h): _artifact(root, path) if path else {"status": "not_configured"}
                                  for h, (path, _) in forecasts.items()},
        },
        "generation_2026": {
            "parsed_rows": len(plants), "plant_count": len(plant_reports),
            "providers": {
                name: {"normalized_2026_rows": sum(p["normalized_source_rows"] for p in plant_reports if p["company"] == name),
                       "profiles": sum(p["company"] == name for p in plant_reports)}
                for name in ("koen", "ewp", "kospo")
            },
            "plants": plant_reports,
        },
        "candidate": _materialized_candidate(
            root, candidate_manifest, baseline, pd.concat(passing) if passing else pd.DataFrame(), plant_reports,
        ),
    }
