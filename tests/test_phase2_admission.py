import json
from pathlib import Path

import pandas as pd
import pytest

from solar_forecast.datasets import phase2_admission as admission
from solar_forecast.datasets.archive_standardizer import HistoricalGenerationStandardizationService
from solar_forecast.datasets.gold_audit import audit_gold_dataset
from solar_forecast.features.future_weather import future_weather_feature_columns
from solar_forecast.infrastructure.artifact_store import sha256_file


def test_retained_kospo_2026_source_hours_and_daily_totals():
    root = Path(__file__).resolve().parents[1]
    reader = HistoricalGenerationStandardizationService(root / admission.GENERATION_ARCHIVE_ROOT, root)
    for source in admission.retained_generation_sources():
        evidence, hourly = admission._read_generation(root, source, reader)
        assert len(hourly) == 1416
        assert hourly["timestamp"].nunique() == 1416
        assert hourly["timestamp"].min() == pd.Timestamp("2026-01-01T00:00:00")
        assert hourly["timestamp"].max() == pd.Timestamp("2026-02-28T23:00:00")
        assert evidence["daily_total_check"]["comparable_days"] == 59
        assert evidence["daily_total_check"]["mismatch_days"] == 0
        assert not hourly["_interval_explicit"].any()


def test_p0_source_inventory_requires_full_bytes_to_match(tmp_path):
    file = tmp_path / "observed.csv"
    file.write_text("original", encoding="utf-8")
    report = tmp_path / "source.json"
    report.write_text(json.dumps({
        "contract": "solar-source-completeness.v1",
        "providers": {"kospo": {"files": [{
            "path": file.name, "sha256": sha256_file(file),
            "first_observed_hour": "2025-01-01T00:00:00", "last_observed_hour": "2026-01-01T00:00:00",
        }, {"path": "older.csv", "last_observed_hour": "2025-12-31T23:00:00"}]}},
    }), encoding="utf-8")
    assert admission._indexed_generation(tmp_path, report) == (admission.GenerationSource("kospo", Path(file.name)),)
    file.write_text("modified", encoding="utf-8")
    with pytest.raises(ValueError, match="differs from the reused collection evidence"):
        admission._indexed_generation(tmp_path, report)


def test_forecast_nominal_origin_does_not_certify_issue_or_receipt(tmp_path):
    archive = tmp_path / "archive.csv"
    expected = pd.date_range("2026-01-02", periods=2, freq="h")
    records = pd.DataFrame({
        "plant_id": ["provider:plant"] * 2,
        "timestamp": expected,
        "horizon_hours": [24, 24],
        "forecast_origin": expected - pd.Timedelta(hours=24),
        "weather_query_latitude": [33.0, 33.0],
        "weather_query_longitude": [126.0, 126.0],
        "grid_latitude": [33.0, 33.0],
        "grid_longitude": [126.0, 126.0],
        "coordinate_source": ["plant_registry_coordinates"] * 2,
        "cell_selection": ["nearest"] * 2,
        "forecast_source": ["retained-provider-log"] * 2,
        "forecast_model": ["provider-model"] * 2,
        **{feature: [1.0, 2.0] for feature in future_weather_feature_columns()},
    })
    records.to_csv(archive, index=False)
    result = admission._forecast(archive, expected, "provider:plant", 24, "aligned_core", 0.98)
    assert result["aligned_targets"] == 2
    assert result["verified_timing_targets"] == 0
    assert result["status"] == "actual_issue_or_receipt_missing"

    records["forecast_issued_at"] = ["2025-12-31T23:00:00+09:00", "2026-01-01T00:00:00+09:00"]
    records["forecast_received_at"] = ["2025-12-31T15:00:00Z", "2026-01-01T01:30:00+09:00"]
    records.to_csv(archive, index=False)
    result = admission._forecast(archive, expected, "provider:plant", 24, "aligned_core", 0.98)
    assert result["verified_timing_targets"] == 1
    assert result["status"] == "insufficient_as_of_coverage"

    records.loc[1, "forecast_received_at"] = "2026-01-01T00:30:00+09:00"
    records.to_csv(archive, index=False)
    assert admission._forecast(archive, expected, "provider:plant", 24, "aligned_core", 0.98)["status"] == "verified"


def test_actual_availability_needs_reconciled_evidence_and_offset(tmp_path):
    retained = tmp_path / "source_receipt.csv"
    retained.write_text("retained source bytes", encoding="utf-8")
    verified = admission._review_evidence(tmp_path, {"evidence_sources": {
        "verified": {"path": retained.name, "sha256": sha256_file(retained)},
        "unverified": {"path": retained.name, "sha256": "stale-or-invalid"},
    }})
    assert list(verified) == ["verified"]
    expected = pd.date_range("2026-01-01", periods=2, freq="h")
    availability = pd.DataFrame({
        "plant_id": ["provider:plant"] * 2,
        "timestamp": expected,
        "generation_available_at": ["2026-01-01T00:59:59+09:00", "2026-01-01T02:03:00+09:00"],
        "weather_available_at": ["2026-01-01T00:10:00", "2026-01-01T01:10:00+09:00"],
        "generation_evidence": ["verified", "unverified"],
        "weather_evidence": ["verified"] * 2,
    })
    coverage = admission._observation_coverage(availability, "provider:plant", expected, set(verified))
    assert coverage["generation_timestamp_rows"] == 0
    assert coverage["weather_timestamp_rows"] == 1
    assert coverage["status"] == "incomplete_actual_observation_availability"


def _write_gold(folder, dates, generation):
    frame = pd.DataFrame({
        "plant_id": ["provider:plant"] * len(dates), "timestamp": dates,
        "company": ["provider"] * len(dates), "energy_source": ["solar"] * len(dates),
        "quality_train_eligible": ["true"] * len(dates),
        "station_id": [185] * len(dates), "generation_mwh": generation,
    })
    folder.mkdir(exist_ok=True)
    dataset = folder / "model_ready.csv.gz"
    frame.to_csv(dataset, index=False)
    parts = folder / "model_ready_parts"
    files = []
    for year, rows in frame.groupby(frame["timestamp"].str[:4]):
        path = parts / "company=provider" / f"year={year}" / "part-000.csv.gz"
        path.parent.mkdir(parents=True, exist_ok=True)
        rows.to_csv(path, index=False)
        files.append({"path": str(path), "rows": len(rows), "bytes": path.stat().st_size,
                      "company": "provider", "year": int(year)})
    manifest = folder / "model_ready_manifest.json"
    manifest.write_text(json.dumps({
        "contract": "solar-model-ready-manifest.v1", "schema_version": 1,
        "dataset": str(dataset), "partitioned_dataset": {"directory": str(parts), "files": files},
        "rows": len(frame), "plants": 1, "start": min(dates), "end": max(dates),
        "target": "generation_mwh", "features": ["station_id"],
        "training_eligibility": {"training_eligible_rows": len(frame)},
    }), encoding="utf-8")
    return dataset, manifest


def test_versioned_candidate_preserves_baseline_and_joins_source(tmp_path, monkeypatch):
    monkeypatch.setattr(admission, "BASELINE_ROWS", 1)
    old, new = "2025-12-31T23:00:00", "2026-01-01T00:00:00"
    dataset, _ = _write_gold(tmp_path / "baseline", [old], [0.25])
    candidate_dir = tmp_path / admission.DATA_VERSION
    _, manifest = _write_gold(candidate_dir, [old, new], [0.25, 0.5])
    expected = pd.DataFrame({
        "plant_id": ["provider:plant"], "timestamp": [pd.Timestamp(new)],
        "generation_mwh": [0.5], "station_id": [185],
    })
    baseline = {"status": "verified", "verified_dataset": str(dataset), "combined_sha256": sha256_file(dataset)}
    report = admission._materialized_candidate(tmp_path, manifest, baseline, expected, [])
    assert report["status"] == "admitted"
    assert report["admitted_2026_rows"] == 1
    assert report["eligible_periods"] == [{"plant_id": "provider:plant", "start": new, "end": new, "rows": 1}]

    _write_gold(candidate_dir, [old, new], [0.26, 0.5])
    with pytest.raises(ValueError, match="changes the verified baseline rows"):
        admission._materialized_candidate(tmp_path, manifest, baseline, expected, [])
    assert sha256_file(dataset) == baseline["combined_sha256"]


def test_stored_gold_receipt_stops_reuse_on_changed_bytes(tmp_path):
    dataset, manifest = _write_gold(tmp_path / "baseline", ["2025-12-31T23:00:00"], [0.25])
    receipt = tmp_path / "verified.json"
    receipt.write_text(json.dumps(audit_gold_dataset(manifest, project_root=tmp_path)), encoding="utf-8")
    assert admission._reuse_verified_gold(tmp_path, manifest, receipt)["status"] == "PASS"
    original = dataset.read_bytes()
    dataset.write_bytes(original[:-1] + bytes([original[-1] ^ 1]))
    assert admission._reuse_verified_gold(tmp_path, manifest, receipt) is None
