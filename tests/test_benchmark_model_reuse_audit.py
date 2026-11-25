"""Ensure frozen-run inventory preserves local evidence and detects identity changes."""

from contextlib import closing
import json
from pathlib import Path
import sqlite3

import pytest

from solar_forecast.config_loader import ModelJobConfig
from solar_forecast.models.shared.checkpoint_store import (
    CHECKPOINT_CONTRACT, dataset_signature, stable_signature, training_fingerprint,
)
from tools.audit_benchmark_model_reuse import artifact_path, inventory_frozen_run, study_identity


def write_json(path: Path, values: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(values), encoding="utf-8")


def test_frozen_inventory_checks_gold_and_legacy_study_without_opening_checkpoint(tmp_path):
    gold = tmp_path / "gold.csv"
    gold.write_text("generation_mwh\n1.0\n", encoding="utf-8")
    benchmark = tmp_path / "artifacts" / "benchmarks" / "frozen"
    run = benchmark / "candidate"
    config_path = run / "resolved_config.json"
    db = run / "optimization.db"
    values = {
        "input_dataset": str(gold), "forecast_horizon_hours": 1,
        "optimizer": {"storage_path": str(db), "study_name": "solar"},
    }
    config = ModelJobConfig("xgboost", "historical_optimized", values, config_path)
    fingerprint = training_fingerprint(config)
    study = study_identity(config, fingerprint)
    signature = stable_signature({"stage": "fixture"})
    stage = f"final_fit_{signature[:20]}"
    payload = run / ".checkpoints" / config.model / fingerprint / f"{stage}.json"
    write_json(payload.with_suffix(".json.meta.json"), {
        "checkpoint_contract": CHECKPOINT_CONTRACT,
        "model": config.model, "fingerprint": fingerprint,
        "stage": stage, "signature": signature,
    })
    payload.write_bytes(b"inventory never deserializes checkpoint bytes")
    write_json(config_path, values)
    write_json(run / "optimization_summary.json", {"study_name": study, "storage_path": str(db)})
    write_json(run / "manifest.json", {
        "model": config.model, "status": "completed",
        "details": {
            "checkpoint": {"stage": stage, "fingerprint": fingerprint},
            "optimizer": {"summary_path": "optimization_summary.json"},
        },
    })
    write_json(benchmark / "manifest.json", {"provenance": {"dataset_fingerprint": dataset_signature(gold)}})
    with closing(sqlite3.connect(db)) as connection:
        connection.execute("CREATE TABLE studies (study_id INTEGER, study_name TEXT)")
        connection.execute("CREATE TABLE trials (trial_id INTEGER, study_id INTEGER)")
        connection.execute("INSERT INTO studies VALUES (1, ?)", (study,))
        connection.execute("INSERT INTO trials VALUES (1, 1)")
        connection.commit()
    original = db.read_bytes()

    report = inventory_frozen_run(benchmark, tmp_path)
    entry, = report["models"]
    assert entry["training_fingerprint_matches"]
    assert entry["optuna_study_name_matches"]
    assert entry["sqlite_exact_study_locations"] == [str(db)]
    stored, = [item for item in entry["stage_inventory"] if item["payload_present"]]
    assert stored["sidecar_identity_matches"]
    assert stored["sidecar_signature_matches_stage_prefix"]
    assert report["checkpoint_payloads_loaded"] is False
    assert db.read_bytes() == original

    gold.write_text("generation_mwh\n2.0\n", encoding="utf-8")
    entry, = inventory_frozen_run(benchmark, tmp_path)["models"]
    assert entry["training_fingerprint_matches"] is False
    assert entry["optuna_study_name_matches"] is False
    assert entry["sqlite_exact_study_locations"] == []


def test_archived_windows_references_require_a_contained_artifact(tmp_path):
    benchmark = tmp_path / "historical"
    path = benchmark / "candidate" / "model.json"
    write_json(path, {})
    assert artifact_path(benchmark, r"C:\ESG_Project1\artifacts\historical\candidate\model.json") == path
    with pytest.raises(ValueError, match="external benchmark artifact"):
        artifact_path(benchmark, "../outside")
