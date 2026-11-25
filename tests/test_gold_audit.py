"""Read-only Gold boundary tests; synthetic CSV and metadata, never model fitting."""
import csv
import gzip
from hashlib import sha256
import json
from pathlib import Path

import pytest

from solar_forecast.config_loader import ModelJobConfig
from solar_forecast.datasets.gold_audit import audit_gold_dataset
from solar_forecast.models.shared.checkpoint_audit import audit_training_compatibility
from solar_forecast.models.shared.checkpoint_store import CHECKPOINT_CONTRACT, training_fingerprint


COLUMNS = ["plant_id", "timestamp", "company", "region", "plant", "energy_source",
           "quality_train_eligible", "f0", "f1", "generation_mwh"]


def _write(path, rows, columns=COLUMNS):
    path.parent.mkdir(parents=True, exist_ok=True)
    opener = gzip.open if path.suffix == ".gz" else open
    with opener(path, "wt", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(columns)
        writer.writerows(rows)


@pytest.fixture
def gold(tmp_path):
    rows = [[f"a:{plant}", f"2024-01-01 {hour:02}:00:00", "a", "제주", plant,
             "solar", "True", str(hour), "", str(hour / 2)]
            for plant in ["신풍리", "용수리"] for hour in range(3)]
    dataset = tmp_path / "gold.csv.gz"
    part = tmp_path / "parts/company=a/year=2024/part.csv.gz"
    _write(dataset, rows)
    # A different export order must not change equality or cardinality.
    _write(part, rows[::-1])
    manifest = {"contract": "solar-model-ready-manifest.v1", "schema_version": 1,
                "dataset": str(dataset), "partitioned_dataset": {"directory": str(tmp_path / "parts"),
                    "files": [{"path": str(part), "company": "a", "year": 2024,
                               "rows": len(rows), "bytes": part.stat().st_size}]},
                "target": "generation_mwh", "features": ["f0", "f1"],
                "rows": 6, "plants": 2, "start": "2024-01-01T00:00:00", "end": "2024-01-01T02:00:00",
                "preprocessing_contract": "solar-observed-preprocessing.v2",
                "training_eligibility": {"training_eligible_rows": 6, "by_energy_source": {
                    "solar": {"retained_rows": 6, "training_eligible_rows": 6}}}}
    mp = tmp_path / "model_ready_manifest.json"
    mp.write_text(json.dumps(manifest), encoding="utf-8")
    return tmp_path, mp, manifest, dataset, part, rows


def _audit(gold):
    root, path, *_ = gold
    return audit_gold_dataset(path, project_root=root, batch_rows=2)


def _update_manifest(gold):
    _, mp, manifest, _, part, _ = gold
    manifest["partitioned_dataset"]["files"][0]["bytes"] = part.stat().st_size
    mp.write_text(json.dumps(manifest), encoding="utf-8")


def _codes(report):
    return {item["code"] for item in report["issues"]}


def test_all_fields_keys_order_independent_and_original_bytes_untouched(gold):
    root = gold[0]
    before = {p: (sha256(p.read_bytes()).hexdigest(), p.stat().st_mtime_ns)
              for p in root.rglob("*") if p.is_file()}
    report = _audit(gold)
    assert report["status"] == "PASS"
    assert report["plants"] == 2
    assert report["dataset"]["rows"] == 6
    assert report["dataset"]["numeric_missing"] == {"f1": 6}
    assert set(report["comparison"].values()) == {0}
    assert report["model_performance_verified"] is False
    assert report["manifest"]["build_time_content_hashes_available"] is False
    assert before == {p: (sha256(p.read_bytes()).hexdigest(), p.stat().st_mtime_ns) for p in before}
    assert set(p for p in root.rglob("*") if p.is_file()) == set(before)


def test_same_keys_same_rows_but_changed_feature_detected(gold):
    rows = [list(row) for row in gold[-1]]
    rows[0][-3] = "9"
    _write(gold[4], rows)
    _update_manifest(gold)
    report = _audit(gold)
    assert "export_content_mismatch" in _codes(report)
    assert report["comparison"]["field_mismatch_rows"] == 1


def test_duplicate_across_insert_batches_is_not_hidden(gold):
    rows = gold[-1]
    _write(gold[4], [*rows, rows[0]])
    _update_manifest(gold)
    report = _audit(gold)
    assert "duplicate_plant_timestamp" in _codes(report)
    assert report["partitions"][0]["duplicate_keys"] == 1
    assert "partition_total_rows_mismatch" in _codes(report)


@pytest.mark.parametrize("field,value,code", [
    ("rows", 99, "manifest_rows_mismatch"),
    ("plants", 99, "manifest_plants_mismatch"),
    ("end", "2024-01-01T03:00:00", "manifest_end_mismatch"),
])
def test_manifest_claims_checked_against_actual(gold, field, value, code):
    gold[2][field] = value
    _update_manifest(gold)
    assert code in _codes(_audit(gold))


def test_partition_row_byte_and_label_claims_checked(gold):
    m = gold[2]["partitioned_dataset"]["files"][0]
    m.update(rows=7, bytes=1, company="other")
    gold[1].write_text(json.dumps(gold[2]), encoding="utf-8")
    assert {"partition_bytes_mismatch", "partition_rows_mismatch", "manifest_partition_label_mismatch"} <= _codes(_audit(gold))


def test_unlisted_partition_cannot_silently_enter_training(gold):
    _write(gold[0] / "parts/company=a/year=2024/unlisted.csv", gold[-1][:1])
    assert "partition_inventory_mismatch" in _codes(_audit(gold))


def test_declared_energy_population_is_checked_from_actual_keys(gold):
    gold[2]["training_eligibility"]["by_energy_source"]["solar"]["retained_plants"] = 99
    _update_manifest(gold)
    result = _audit(gold)
    assert "manifest_energy_population_mismatch" in _codes(result)
    assert result["population_by_energy_source"]["solar"]["retained_plants"] == 2


def test_negative_target_cannot_be_marked_training_eligible(gold):
    rows = [list(row) for row in gold[-1]]
    rows[0][-1] = "-1"
    _write(gold[3], rows)
    _write(gold[4], rows)
    _update_manifest(gold)
    assert "invalid_eligible_target" in _codes(_audit(gold))


def test_nonfinite_features_are_reported_without_rewriting_frozen_missing_policy(gold):
    rows = [list(row) for row in gold[-1]]
    rows[0][-3] = "inf"
    _write(gold[3], rows)
    _write(gold[4], rows)
    _update_manifest(gold)
    result = _audit(gold)
    assert result["status"] == "PASS"
    assert result["dataset"]["numeric_nonfinite"] == {"f0": 1}
    with gzip.open(gold[3], "rt", encoding="utf-8-sig") as stream:
        assert list(csv.reader(stream))[1][-3] == "inf"


def test_column_reordering_is_detected(gold):
    rows = [row[:-3] + [row[-2], row[-3], row[-1]] for row in gold[-1]]
    columns = COLUMNS[:-3] + ["f1", "f0", "generation_mwh"]
    _write(gold[4], rows, columns)
    _update_manifest(gold)
    assert "partition_schema_order_mismatch" in _codes(_audit(gold))


@pytest.mark.parametrize("timestamp", ["not-a-time", "2024-01-01T00:00:00+09:00", "2024-01-01T00:30:00"])
def test_invalid_aware_and_nonhourly_timestamps_rejected_without_conversion(gold, timestamp):
    rows = [list(row) for row in gold[-1]]
    rows[0][1] = timestamp
    _write(gold[3], rows)
    with pytest.raises(ValueError, match="timestamp"):
        _audit(gold)


def test_bad_column_count_fails_clearly(gold):
    _write(gold[4], [gold[-1][0][:-1]])
    with pytest.raises(ValueError, match="width"):
        _audit(gold)


def test_partition_path_escape_rejected_before_read(gold):
    gold[2]["partitioned_dataset"]["files"][0]["path"] = str(gold[0].parent / "outside.csv")
    gold[1].write_text(json.dumps(gold[2]), encoding="utf-8")
    with pytest.raises(ValueError, match="within the project"):
        _audit(gold)


def test_external_path_rejected_before_filesystem_resolution(gold, monkeypatch):
    from solar_forecast.datasets.gold_audit import _resolve
    root = gold[0].resolve()
    external = root.parent / "outside.csv"
    original = Path.resolve

    def guarded(path, *args, **kwargs):
        if path == external:
            raise AssertionError("External path must not be probed")
        return original(path, *args, **kwargs)

    monkeypatch.setattr(Path, "resolve", guarded)
    with pytest.raises(ValueError, match="within the project"):
        _resolve(root, str(external))
    with pytest.raises(ValueError, match="not URLs"):
        _resolve(root, "https://example.invalid/data.csv")


def _config(gold):
    values = {"model": "cnn_bilstm", "profile": "test", "input_dataset": str(gold[0] / "parts"),
              "target_column": "generation_mwh", "feature_columns": ["f1", "f0"],
              "quality_filter_column": "quality_train_eligible", "energy_source_filter": "solar",
              "checkpoint": {"root": str(gold[0] / "checkpoints")}}
    path = gold[0] / "config.json"
    path.write_text(json.dumps(values), encoding="utf-8")
    return ModelJobConfig("cnn_bilstm", "test", values, path)


def test_no_checkpoint_is_not_a_false_resume_pass(gold):
    report = audit_training_compatibility(_config(gold), _audit(gold), project_root=gold[0])
    assert report["schema_status"] == "PASS"
    assert report["ordered_features"] == ["f1", "f0"]
    assert report["checkpoint_status"] == "NO_EXACT_TRAINING_IDENTITY"
    assert report["resume_executed"] is False
    assert not (gold[0] / "checkpoints").exists()


def test_relative_path_spelling_is_not_rewritten_in_existing_fingerprint(gold, monkeypatch):
    from solar_forecast.models.shared import checkpoint_store
    cfg = _config(gold)
    cfg.values["input_dataset"] = "parts"
    monkeypatch.setattr(checkpoint_store, "PROJECT_ROOT", gold[0])
    expected = training_fingerprint(cfg)
    report = audit_training_compatibility(cfg, _audit(gold), project_root=gold[0])
    assert report["training_fingerprint"] == expected
    assert cfg.values["input_dataset"] == "parts"


def test_metadata_matches_but_never_claims_stage_or_payload_resume(gold):
    cfg = _config(gold)
    fingerprint = training_fingerprint(cfg)
    folder = gold[0] / "checkpoints/cnn_bilstm" / fingerprint
    folder.mkdir(parents=True)
    (folder / "trial.pt").write_bytes(b"not-deserialized")
    meta = {"checkpoint_contract": CHECKPOINT_CONTRACT, "model": "cnn_bilstm", "fingerprint": fingerprint,
            "stage": "trial", "signature": "specific-stage", "completed": False,
            "progress": {"next_epoch": 13, "total_epochs": 20}}
    (folder / "trial.pt.meta.json").write_text(json.dumps(meta), encoding="utf-8")
    report = audit_training_compatibility(cfg, _audit(gold), project_root=gold[0])
    assert report["training_fingerprint"] == fingerprint
    assert report["matching_metadata_count"] == 1
    assert report["checkpoint_status"] == "MATCHING_METADATA_REQUIRES_STAGE_SIGNATURE"
    assert report["checkpoint_payload_loaded"] is False
    cfg.values["forecast_horizon_hours"] = 72
    assert audit_training_compatibility(cfg, _audit(gold), project_root=gold[0])["matching_metadata_count"] == 0


def test_config_missing_feature_and_wrong_data_rejected(gold):
    cfg = _config(gold)
    cfg.values["feature_columns"].append("not_in_gold")
    report = audit_training_compatibility(cfg, _audit(gold), project_root=gold[0])
    assert report["schema_status"] == "FAIL"
    assert report["missing_columns"] == ["not_in_gold"]
    other = gold[0] / "other.csv"
    _write(other, gold[-1])
    cfg.values["input_dataset"] = str(other)
    report = audit_training_compatibility(cfg, _audit(gold), project_root=gold[0])
    assert "config_does_not_reference_audited_gold" in report["schema_issues"]


def test_cli_parser_dispatches_audit_without_running_a_job():
    from solar_forecast.cli.parser import build_parser
    args = build_parser().parse_args(["audit-gold", "--batch-rows", "3", "--config", "config/models/xgboost.json"])
    assert args.func.__name__ == "handle_audit_gold_command"
    assert args.batch_rows == 3
    assert args.config == ["config/models/xgboost.json"]


@pytest.mark.parametrize("as_json", [False, True])
def test_actual_cli_handler_reports_without_training(gold, monkeypatch, capsys, as_json):
    from argparse import Namespace
    from solar_forecast import config_loader
    from solar_forecast.cli.dataset_commands import handle_audit_gold_command
    cfg = _config(gold)
    monkeypatch.setattr(config_loader, "PROJECT_ROOT", gold[0])
    handle_audit_gold_command(Namespace(manifest=str(gold[1]), config=[str(cfg.source)], batch_rows=2, json=as_json))
    output = capsys.readouterr().out
    if as_json:
        result = json.loads(output)
        assert result["status"] == "PASS"
        assert result["training_performed"] is False
        assert result["training_compatibility"][0]["resume_executed"] is False
    else:
        assert "Gold/manifest consistency: PASS" in output
        assert "NO_EXACT_TRAINING_IDENTITY" in output


def test_cli_handler_fails_nonzero_on_real_data_difference(gold, monkeypatch, capsys):
    from argparse import Namespace
    from solar_forecast import config_loader
    from solar_forecast.cli.dataset_commands import handle_audit_gold_command
    cfg = _config(gold)
    monkeypatch.setattr(config_loader, "PROJECT_ROOT", gold[0])
    _write(gold[4], gold[-1][1:])
    _update_manifest(gold)
    with pytest.raises(SystemExit) as exc:
        handle_audit_gold_command(Namespace(manifest=str(gold[1]), config=[str(cfg.source)], batch_rows=2, json=True))
    assert exc.value.code == 2
    assert json.loads(capsys.readouterr().out)["status"] == "FAIL"


def test_checkpoint_cli_reads_metadata_without_gold_manifest_or_model_load(gold, monkeypatch, capsys):
    from solar_forecast.cli.parser import build_parser
    from solar_forecast.models.shared.checkpoint_store import TrainingCheckpointStore

    cfg = _config(gold)
    gold[1].unlink()

    def forbidden(*args, **kwargs):
        pytest.fail("Checkpoint audit must not deserialize a model")

    monkeypatch.setattr(TrainingCheckpointStore, "load_torch", forbidden)
    monkeypatch.setattr(TrainingCheckpointStore, "load_xgboost", forbidden)
    args = build_parser().parse_args(["checkpoint-audit", "--config", str(cfg.source), "--json"])
    args.func(args)
    result = json.loads(capsys.readouterr().out)
    assert result["checkpoint_scope"] == "configured_fingerprint_and_saved_metadata_only"
    assert result["checkpoint_status"] == "NO_EXACT_TRAINING_IDENTITY"
    assert result["checkpoint_payload_loaded"] is False
    assert result["resume_executed"] is False


def test_gold_consistency_pass_is_reported_independently_of_model_schema_failure(gold, monkeypatch, capsys):
    from solar_forecast import config_loader
    from solar_forecast.cli.parser import build_parser

    cfg = _config(gold)
    cfg.values["feature_columns"].append("absent_from_gold")
    cfg.source.write_text(json.dumps(cfg.values), encoding="utf-8")
    monkeypatch.setattr(config_loader, "PROJECT_ROOT", gold[0])
    args = build_parser().parse_args(["audit-gold", "--manifest", str(gold[1]),
                                      "--config", str(cfg.source), "--json"])
    with pytest.raises(SystemExit) as error:
        args.func(args)
    assert error.value.code == 2
    result = json.loads(capsys.readouterr().out)
    assert result["gold_consistency_status"] == "PASS"
    assert result["status"] == "FAIL"
    assert result["training_compatibility"][0]["schema_status"] == "FAIL"
    assert result["data_admission_checked"] is False
