from pathlib import Path
import json
from types import SimpleNamespace

import pytest

from solar_forecast.jobs.lock import TrainingAlreadyRunning, exclusive_training_lock
from solar_forecast.jobs import training
from solar_forecast.settings import ModelJobConfig


def test_training_lock_blocks_second_model(tmp_path):
    lock = tmp_path / ".training.lock"
    with exclusive_training_lock(lock, "xgboost"):
        with pytest.raises(TrainingAlreadyRunning):
            with exclusive_training_lock(lock, "cnn_bilstm"):
                pass
    assert not lock.exists()


def test_completed_training_manifest_preserves_execution_mode(tmp_path, monkeypatch):
    monkeypatch.setattr(training, "PROJECT_ROOT", tmp_path)
    monkeypatch.setattr(
        training.importlib,
        "import_module",
        lambda _: SimpleNamespace(train=lambda config, run_dir, smoke: {"metrics": {}}),
    )
    config = ModelJobConfig(
        model="xgboost",
        profile="test",
        values={
            "output_root": "artifacts/models/xgboost",
            "energy_source_filter": "solar",
            "target_column": "generation_mwh",
            "forecast_horizon_hours": 24,
            "evaluation_protocol": "rolling_origin",
        },
        source=tmp_path / "config.json",
    )

    run_dir = training.TrainingService(
        trainers={"xgboost": ("fake_trainer", "train")}
    ).run(config, smoke=True)
    manifest = json.loads((run_dir / "manifest.json").read_text(encoding="utf-8"))

    assert manifest["status"] == "completed"
    assert manifest["details"]["run_context"] == {
        "execution_mode": "smoke",
        "energy_source": "solar",
        "target": "generation_mwh",
        "target_unit": "MWh",
        "horizon_hours": 24,
        "evaluation_protocol": "rolling_origin",
    }
