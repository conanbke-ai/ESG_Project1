"""Synthetic ownership/CLI contracts: no datasets, model training or delivery."""
from __future__ import annotations

import ast
import json
import os
from pathlib import Path
import subprocess
import sys

import pytest

from solar_forecast.cli import main
from solar_forecast.jobs.contracts import job_contract_catalog, list_job_contracts
from solar_forecast.jobs.storage_layout import (
    STORAGE_LAYOUT_CONTRACT,
    StorageLayer,
    build_storage_layout_catalog,
    classify_storage_path,
    list_storage_locations,
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]


@pytest.mark.parametrize("path,storage_id,layer", [
    ("file/solar_data_file/provider/raw.csv", "generation_archive", StorageLayer.BRONZE),
    ("file/KMA_data_file/OBS_ASOS_TIM_2025.csv", "weather_archive", StorageLayer.BRONZE),
    ("file/raw/source/run.csv", "bronze", StorageLayer.BRONZE),
    ("file/standardized/downloads/source.csv", "collector_silver", StorageLayer.SILVER),
    ("file/standardized/generation/part.csv.gz", "silver", StorageLayer.SILVER),
    ("file/standardized/plant_registry.csv", "plant_registry", StorageLayer.SILVER),
    ("file/standardized/model_ready.csv.gz", "gold_dataset", StorageLayer.GOLD),
    ("file/standardized/model_ready_parts/company/part.csv.gz", "gold_partitions", StorageLayer.GOLD),
    ("file/standardized/model_ready_manifest.json", "gold_manifest", StorageLayer.GOLD),
    ("file/merge_data/val.csv", "legacy_merged", StorageLayer.LEGACY),
    ("data/json/managed_human.json", "reference_data", StorageLayer.REFERENCE),
    ("artifacts/checkpoints/cnn_bilstm/fingerprint/latest.pt", "checkpoints", StorageLayer.ARTIFACTS),
    ("artifacts/optimization/study.db", "optuna", StorageLayer.ARTIFACTS),
    ("artifacts/models/xgboost/run/model.json", "model_runs", StorageLayer.ARTIFACTS),
    ("artifacts/benchmarks/run/manifest.json", "artifacts", StorageLayer.ARTIFACTS),
    ("dashboard/data/dashboard_data.json", "dashboard", StorageLayer.REPORTING),
    ("map/json/geoJson.json", "map", StorageLayer.REFERENCE),
    ("map/html/solar_dashboard.html", "map_redirects", StorageLayer.REPORTING),
    ("solar_dashboard.html", "redirect_solar_dashboard", StorageLayer.REPORTING),
    ("src/solar_forecast/models/cnn_bilstm/trainer.py", "source", StorageLayer.SOURCE),
    ("config/models/xgboost.json", "config", StorageLayer.CONFIG),
])
def test_declared_paths_preserve_existing_physical_roles(tmp_path, path, storage_id, layer):
    query = classify_storage_path(tmp_path, path)
    assert query["location"]["storage_id"] == storage_id
    assert query["location"]["layer"] == layer
    assert query["path"] == path
    assert query["data_integrity_checked"] is False
    assert query["model_eligibility_checked"] is False
    assert not list(tmp_path.iterdir())


def test_catalog_is_deterministic_and_has_unique_location_ids(tmp_path):
    first = build_storage_layout_catalog(tmp_path)
    assert first == build_storage_layout_catalog(tmp_path)
    assert first["contract"] == STORAGE_LAYOUT_CONTRACT
    assert first["read_only"] is True
    assert first["automatic_move_allowed"] is False
    assert first["automatic_delete_allowed"] is False
    assert first["write_protection_enforced"] is False
    assert first["query"] is None
    ids = [entry["storage_id"] for entry in first["locations"]]
    assert len(ids) == len(set(ids))
    assert all(not Path(entry["path"]).is_absolute() for entry in first["locations"])


def test_job_input_output_references_reuse_exact_existing_contracts(tmp_path):
    before = json.dumps(job_contract_catalog(), sort_keys=True)
    references = build_storage_layout_catalog(tmp_path)["job_artifact_references"]
    expected = {
        (job.job_id, direction, artifact.name, artifact.contract, artifact.path_pattern)
        for job in list_job_contracts()
        for direction, artifacts in (("input", job.inputs), ("output", job.outputs))
        for artifact in artifacts
    }
    actual = {(r["job_id"], r["direction"], r["artifact"], r["contract"], r["path_pattern"]) for r in references}
    assert actual == expected
    assert all(r["storage_id"] is not None for r in references)
    assert json.dumps(job_contract_catalog(), sort_keys=True) == before


@pytest.mark.parametrize("path", ["file/standardized/model_ready_parts_old/data.csv", "file/standardized/model_ready.csv.gz.backup"])
def test_gold_matching_does_not_use_unsafe_string_prefixes(tmp_path, path):
    query = classify_storage_path(tmp_path, path)
    assert query["location"]["storage_id"] == "silver"


def test_unknown_path_is_not_promoted_to_gold(tmp_path):
    query = classify_storage_path(tmp_path, "unreviewed/new.csv")
    assert query["classification"] == "unclassified_review_required"
    assert query["location"] is None


def test_declared_file_cannot_be_interpreted_as_a_directory(tmp_path):
    with pytest.raises(ValueError, match="declared file"):
        classify_storage_path(tmp_path, "file/standardized/model_ready.csv.gz/child")


@pytest.mark.parametrize("path", ["../outside.csv", "file/../../outside.csv", "https://example.invalid/data.csv"])
def test_escape_and_remote_paths_are_rejected(tmp_path, path):
    with pytest.raises(ValueError):
        classify_storage_path(tmp_path, path)


def test_absolute_contained_path_is_reported_without_host_prefix(tmp_path):
    query = classify_storage_path(tmp_path, tmp_path / "artifacts/checkpoints/a.pt")
    assert query["path"] == "artifacts/checkpoints/a.pt"
    assert str(tmp_path) not in json.dumps(query)


def test_lexically_external_path_is_rejected_before_candidate_resolution(tmp_path, monkeypatch):
    original = Path.resolve
    resolved_paths = []
    def resolve(path, *args, **kwargs):
        resolved_paths.append(path)
        return original(path, *args, **kwargs)
    monkeypatch.setattr(Path, "resolve", resolve)
    with pytest.raises(ValueError, match="outside"):
        classify_storage_path(tmp_path, tmp_path.parent / "external" / "input.csv")
    assert resolved_paths == [tmp_path]


def test_symlink_cannot_claim_ownership_of_an_external_file(tmp_path):
    outside = tmp_path.parent / (tmp_path.name + "-external")
    outside.mkdir()
    link = tmp_path / "file"
    try:
        link.symlink_to(outside, target_is_directory=True)
    except OSError:
        pytest.skip("Symlink creation not permitted by this test environment")
    with pytest.raises(ValueError, match="outside"):
        classify_storage_path(tmp_path, "file/private.csv")


def test_resolved_internal_alias_cannot_escape_even_without_symlink_privilege(tmp_path, monkeypatch):
    original = Path.resolve
    candidate = tmp_path / "file" / "alias.csv"
    outside = tmp_path.parent / "external" / "alias.csv"
    def resolve(path, *args, **kwargs):
        return outside if path == candidate else original(path, *args, **kwargs)
    monkeypatch.setattr(Path, "resolve", resolve)
    with pytest.raises(ValueError, match="outside"):
        classify_storage_path(tmp_path, "file/alias.csv")


def test_describing_layout_never_enumerates_or_opens_artifacts(tmp_path, monkeypatch):
    def forbidden(*args, **kwargs):
        raise AssertionError("Storage catalog must not read/enumerate/write artifacts")
    for name in ("open", "read_text", "read_bytes", "write_text", "write_bytes", "mkdir", "glob", "rglob", "iterdir", "rename", "replace", "unlink"):
        monkeypatch.setattr(Path, name, forbidden)
    result = build_storage_layout_catalog(tmp_path, "file/standardized/model_ready.csv.gz")
    assert result["query"]["location"]["layer"] == StorageLayer.GOLD


def test_cli_routes_to_read_only_catalog_without_creating_files(tmp_path, monkeypatch, capsys):
    import solar_forecast.cli.job_commands as commands
    monkeypatch.setattr(commands, "PROJECT_ROOT", tmp_path)
    main(["storage-layout", "--json", "--path", "file/standardized/model_ready_parts/p.csv.gz"])
    payload = json.loads(capsys.readouterr().out)
    assert payload["query"]["location"]["layer"] == "gold"
    assert not list(tmp_path.iterdir())


def test_cli_text_output_labels_scope_and_unknown_query(tmp_path, monkeypatch, capsys):
    import solar_forecast.cli.job_commands as commands
    monkeypatch.setattr(commands, "PROJECT_ROOT", tmp_path)
    main(["storage-layout", "--path", "unreviewed.csv"])
    output = capsys.readouterr().out
    assert "read-only" in output and "no integrity/admission" in output
    assert "unclassified_review_required" in output


def test_cli_rejects_escape_without_creating_output(tmp_path, monkeypatch):
    import solar_forecast.cli.job_commands as commands
    monkeypatch.setattr(commands, "PROJECT_ROOT", tmp_path)
    with pytest.raises(SystemExit, match="outside"):
        main(["storage-layout", "--path", "../private"])
    assert not list(tmp_path.iterdir())


def test_real_cli_catalog_does_not_import_training_frameworks():
    code = """
import sys
from solar_forecast.cli import main
main(['storage-layout', '--json'])
assert not ({'torch','xgboost','optuna','pandas','selenium'} & set(sys.modules))
"""
    env = {**os.environ, "PYTHONPATH": str(PROJECT_ROOT / "src"), "PYTHONDONTWRITEBYTECODE": "1"}
    completed = subprocess.run([sys.executable, "-B", "-c", code], cwd=PROJECT_ROOT, env=env, capture_output=True, text=True, timeout=30)
    assert completed.returncode == 0, completed.stderr
    assert json.loads(completed.stdout)["contract"] == STORAGE_LAYOUT_CONTRACT


def test_layout_module_preserves_dependency_direction_and_no_eager_model_imports():
    source = PROJECT_ROOT / "src/solar_forecast/jobs/storage_layout.py"
    tree = ast.parse(source.read_text(encoding="utf-8"))
    modules = {node.module for node in ast.walk(tree) if isinstance(node, ast.ImportFrom)}
    assert not any(name and name.startswith(("solar_forecast.cli", "solar_forecast.models", "solar_forecast.collectors", "solar_forecast.runtime")) for name in modules)
    assert {location.layer for location in list_storage_locations()} >= {
        StorageLayer.BRONZE, StorageLayer.SILVER, StorageLayer.GOLD,
        StorageLayer.ARTIFACTS, StorageLayer.REPORTING,
    }
