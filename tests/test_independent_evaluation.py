from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path

import pandas as pd
import pytest

from solar_forecast.evaluation.independent_evaluation import (
    ADMISSION_CONTRACT,
    DEVELOPMENT_STATUS,
    EXPOSURE_CONTRACT,
    HORIZONS,
    PREDICTION_KEY,
    run_independent_evaluation_audit,
)
from solar_forecast.infrastructure.artifact_store import sha256_file


CONFIG = Path(__file__).resolve().parents[1] / "config/experiments/optimized.json"


def write(path, record):
    path.write_text(json.dumps(record), encoding="utf-8")
    return path


@pytest.fixture
def development_run(tmp_path):
    directory = tmp_path / "benchmark"
    directory.mkdir()
    write(directory / "manifest.json", {"contract": "solar-optimized-benchmark.v1", "status": "completed",
          "provenance": {"dataset_fingerprint": "fixture", "code_revision": "fixture",
                         "data_provenance": {"evaluation_status": DEVELOPMENT_STATUS}}})
    write(directory / "experiment.json", json.loads(CONFIG.read_text(encoding="utf-8")))
    for horizon in HORIZONS:
        subdir = directory / f"horizon_{horizon}h"
        subdir.mkdir()
        timestamps = pd.date_range("2025-01-08", periods=2, freq="h")
        test = pd.DataFrame({"timestamp": timestamps, "forecast_origin": timestamps - pd.Timedelta(hours=horizon),
                             "horizon_hours": horizon, "plant_id": "koen:test", "plant": "test",
                             "region": "test", "split": "test", "y_true": [0., 2.], "xgb_pred": [0., 2.],
                             "cnn_pred": [1., 3.], "hybrid_pred": [0.25, 2.25], "persistence_pred": [2., 4.]})
        test_path = subdir / "test_predictions.csv.gz"
        test.to_csv(test_path, index=False)
        profiles = subdir / "gate_profiles.csv"
        profiles.write_bytes(b"scope,weight\nglobal,0.5\n")
        write(subdir / "selection.json", {"status": "completed",
              "evaluation_contract": {"horizon_hours": horizon, "target_unit": "MWh", "target": "generation_mwh",
                                      "dataset_fingerprint": "fixture", "prediction_key": PREDICTION_KEY,
                                      "information_set": "observed", "observation_availability_assumption": "fixture"},
              "selection_rule": {"test_used_for_selection": False}, "purge": {"selection_gap_hours": 168},
              "periods": {"test": {"start": timestamps[0].isoformat(), "end": timestamps[1].isoformat(),
                                   "rows": 2, "plants": 1, "unique_timestamps": 2}},
              "test_metrics": {model: {"pooled": {"n_samples": 2, "mae": error, "rmse": error, "r2": 1 - error**2}}
                               for model, error in (("xgboost", 0), ("cnn_bilstm", 1), ("hybrid", 0.25), ("persistence", 2))},
              "files": {"test_predictions": test_path.name, "gate_profiles": profiles.name},
              "files_sha256": {"test_predictions": sha256_file(test_path),
                               "gate_profiles": sha256(b"scope,weight\r\nglobal,0.5\r\n").hexdigest()}})
    return directory


def admission(tmp_path, *, ready=False, duplicate=False):
    candidate = {"status": "blocked", "admitted_2026_rows": 0, "eligible_periods": [], "dataset_path": None}
    if ready:
        snapshot = tmp_path / "candidate.csv"
        snapshot.write_text("unopened holdout fixture", encoding="utf-8")
        period = {"plant_id": "koen:test", "start": "2026-01-01T00:00:00", "end": "2026-01-10T23:00:00", "rows": 240}
        candidate.update(status="admitted", period_end_inclusive=True,
                         dataset_path=str(snapshot), dataset_sha256=sha256_file(snapshot),
                         admitted_2026_rows=480 if duplicate else 240,
                         eligible_periods=[period, period] if duplicate else [period])
    return write(tmp_path / "admission.json", {"contract": ADMISSION_CONTRACT, "data_version": "fixture",
                                               "time_zone": "Asia/Seoul", "horizons_hours": list(HORIZONS),
                                               "candidate": candidate})


def audit(development_run, admission_path, **kwargs):
    return run_independent_evaluation_audit(config_path=CONFIG, development_run=development_run,
                                            admission_path=admission_path, **kwargs)


def test_blocked_holdout_keeps_twelve_comparisons_and_verified_development_evidence(tmp_path, development_run):
    result = audit(development_run, admission(tmp_path))
    assert result["status"] == "blocked_no_admitted_2026_targets"
    assert result["holdout"]["candidate_target_rows"] == 0
    for historical in result["development_evaluation"]["horizons"]:
        assert historical["status"] == DEVELOPMENT_STATUS
        assert historical["gate_profiles_integrity"]["status"] == "crlf_reconstruction_matches"
    tasks = result["comparison_protocol"]["tasks"]
    assert [task["horizon_hours"] for task in tasks] == [1, 24, 72]
    assert [task["forecast_weather_enabled"] for task in tasks] == [False, True, True]
    assert all(task["models"] == ["xgboost", "cnn_bilstm", "hybrid", "persistence"] for task in tasks)
    assert not result["independent_scoring_performed"]


def test_admitted_2026_requires_a_candidate_bound_exposure_review_and_168_hour_purge(tmp_path, development_run):
    candidate = admission(tmp_path, ready=True)
    assert audit(development_run, candidate)["status"] == "blocked_exposure_review_pending"
    source = write(tmp_path / "ledger.json", {"fixture": True})
    review = {"contract": EXPOSURE_CONTRACT, "status": "complete", "data_version": "fixture",
              "admission_manifest_sha256": sha256_file(candidate),
              "reviewed_period": {"start": "2026-01-01T00:00:00", "end": "2026-01-10T23:00:00"},
              "sources": [{"path": source.name, "sha256": sha256_file(source)}],
              "exposed_periods": [{"start": "2026-01-01T00:00:00", "end": "2026-01-02T23:00:00"}]}
    review_path = write(tmp_path / "review.json", review)
    result = audit(development_run, candidate, exposure_path=review_path)
    assert result["holdout"]["verified_candidate_periods"] == [
        {"plant_id": "koen:test", "start": "2026-01-10T00:00:00", "end": "2026-01-10T23:00:00",
         "candidate_target_rows": 24}]
    assert result["holdout"]["prediction_cohort_rows"] is None
    assert not result["holdout_labels_read"]
    source.write_text("changed", encoding="utf-8")
    with pytest.raises(ValueError, match="SHA-256 mismatch"):
        audit(development_run, candidate, exposure_path=review_path)


def test_rejects_overlapping_admission_and_modified_saved_predictions(tmp_path, development_run):
    with pytest.raises(ValueError, match="periods for a plant overlap"):
        audit(development_run, admission(tmp_path, ready=True, duplicate=True))
    (development_run / "horizon_1h/test_predictions.csv.gz").write_bytes(b"changed")
    with pytest.raises(ValueError, match="SHA-256 mismatch"):
        audit(development_run, admission(tmp_path))


def test_audit_output_cannot_replace_or_extend_input_artifacts(tmp_path, development_run):
    candidate = admission(tmp_path)
    review = write(tmp_path / "review.json", {})
    outputs = (candidate, review, CONFIG, development_run / "horizon_1h/test_predictions.csv.gz",
               development_run / "unwritten-report.json")
    for path in outputs:
        original = path.read_bytes() if path.is_file() else None
        with pytest.raises(FileExistsError, match="outside all input artifacts"):
            audit(development_run, candidate, exposure_path=review, output_path=path)
        if original is None:
            assert not path.exists()
        else:
            assert path.read_bytes() == original
