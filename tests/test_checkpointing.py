import json
from pathlib import Path

import numpy as np
import pandas as pd
import pytest
import torch
from xgboost import XGBRegressor

from solar_forecast.models.checkpointing import (
    TrainingCheckpointStore,
    training_fingerprint,
)
from solar_forecast.models.cnn import adaptive, workflow
from solar_forecast.models.cnn.config import SequenceConfig
from solar_forecast.models.cnn.model import ModelConfig
from solar_forecast.models.xgboost_checkpoint import fit_xgboost_resumable
from solar_forecast.settings import ModelJobConfig


def _store(tmp_path: Path, model: str = "cnn_bilstm") -> TrainingCheckpointStore:
    return TrainingCheckpointStore(
        tmp_path / "checkpoints",
        model=model,
        fingerprint="test-fingerprint",
        cnn_every_epochs=1,
        xgboost_every_rounds=1,
    )


def _frame(rows: int = 100) -> pd.DataFrame:
    rng = np.random.default_rng(9)
    return pd.DataFrame(
        {
            "f0": rng.normal(size=rows),
            "f1": rng.normal(size=rows),
            "target": rng.normal(size=rows),
        }
    )


def test_torch_checkpoint_is_atomic_and_rejects_wrong_signature(tmp_path: Path):
    store = _store(tmp_path)
    path = store.save_torch(
        "final",
        {"next_epoch": 2, "tensor": torch.tensor([1.0])},
        signature="compatible",
        progress={"next_epoch": 2},
    )

    assert path is not None and path.exists()
    assert not path.with_name(f"{path.stem}.tmp{path.suffix}").exists()
    assert store.load_torch("final", signature="compatible")["next_epoch"] == 2
    with pytest.raises(ValueError, match="incompatible"):
        store.load_torch("final", signature="different")


def test_training_fingerprint_changes_with_data_but_not_runtime_budget(tmp_path: Path):
    dataset = tmp_path / "dataset.csv"
    dataset.write_text("value\n1\n", encoding="utf-8")
    values = {
        "model": "xgboost",
        "profile": "optimized",
        "input_dataset": str(dataset),
        "feature_columns": ["value"],
        "optimizer": {"study_name": "solar", "max_trials": 10},
        "checkpoint": {"root": str(tmp_path / "first")},
    }
    config = ModelJobConfig("xgboost", "optimized", values, tmp_path / "config.json")
    first = training_fingerprint(config)
    values["optimizer"]["max_trials"] = 20
    values["checkpoint"]["root"] = str(tmp_path / "second")

    assert training_fingerprint(config) == first
    dataset.write_text("value\n2\n", encoding="utf-8")
    assert training_fingerprint(config) != first


def test_cnn_fixed_training_resumes_after_interrupted_epoch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    store = _store(tmp_path)
    original = workflow._train_one_epoch
    calls = 0

    def interrupted(*args, **kwargs):
        nonlocal calls
        calls += 1
        if calls == 2:
            raise RuntimeError("simulated interruption")
        return original(*args, **kwargs)

    monkeypatch.setattr(workflow, "_train_one_epoch", interrupted)
    with pytest.raises(RuntimeError, match="simulated interruption"):
        workflow.train_and_save(
            _frame(),
            "target",
            feature_columns=["f0", "f1"],
            sequence_config=SequenceConfig(
                sequence_length=5,
                val_size=0.2,
                calibration_size=0.1,
                test_size=0.2,
                batch_size=16,
                shuffle=False,
            ),
            epochs=3,
            use_optuna=False,
            output_dir=str(tmp_path / "runs"),
            checkpoint_store=store,
        )

    resumed_calls = 0

    def resumed(*args, **kwargs):
        nonlocal resumed_calls
        resumed_calls += 1
        return original(*args, **kwargs)

    monkeypatch.setattr(workflow, "_train_one_epoch", resumed)
    artifacts = workflow.train_and_save(
        _frame(),
        "target",
        feature_columns=["f0", "f1"],
        sequence_config=SequenceConfig(
            sequence_length=5,
            val_size=0.2,
            calibration_size=0.1,
            test_size=0.2,
            batch_size=16,
            shuffle=False,
        ),
        epochs=3,
        use_optuna=False,
        output_dir=str(tmp_path / "runs"),
        checkpoint_store=store,
    )

    assert resumed_calls == 2
    assert artifacts["checkpoint"]["resumed"] is True
    assert Path(artifacts["checkpoint_path"]).exists()


def test_cnn_legacy_workflow_persists_study_and_final_state(tmp_path: Path):
    arguments = {
        "frame": _frame(120),
        "target_column": "target",
        "feature_columns": ["f0", "f1"],
        "sequence_config": SequenceConfig(
            sequence_length=5,
            val_size=0.2,
            calibration_size=0.1,
            test_size=0.2,
            batch_size=16,
            shuffle=False,
        ),
        "n_trials": 1,
        "output_dir": str(tmp_path / "runs"),
        "use_optuna": True,
        "epochs": 1,
        "optimizer_trial_epochs": 1,
        "early_stopping_patience": 1,
        "checkpoint_root": tmp_path / "checkpoints",
        "optimizer_storage_path": tmp_path / "optimization.db",
    }
    first = workflow.train_and_save(**arguments)
    second = workflow.train_and_save(**arguments)
    first_summary = json.loads(
        Path(first["optimizer"]["summary_path"]).read_text(encoding="utf-8")
    )
    second_summary = json.loads(
        Path(second["optimizer"]["summary_path"]).read_text(encoding="utf-8")
    )

    assert first_summary["executed_trials"] == 1
    assert second_summary["executed_trials"] == 0
    assert second["checkpoint"]["resumed"] is True


def test_adaptive_training_restores_bandit_and_epoch_progress(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    store = _store(tmp_path)
    calls = 0

    def interrupted(*args, **kwargs):
        nonlocal calls
        calls += 1
        if calls == 2:
            raise RuntimeError("simulated adaptive interruption")
        return 0.5

    monkeypatch.setattr(adaptive, "_train_one_epoch", interrupted)
    monkeypatch.setattr(
        adaptive,
        "_evaluate",
        lambda *args, **kwargs: {"loss": 0.4, "mae": 0.3, "rmse": 0.4, "r2": 0.0},
    )
    with pytest.raises(RuntimeError, match="adaptive interruption"):
        adaptive.run_adaptive_training(
            ModelConfig(n_features=2, cnn_channels=4, lstm_hidden=4, dense_units=4),
            train_loader=[],
            val_loader=[],
            epochs=3,
            checkpoint_store=store,
            checkpoint_stage="adaptive",
            checkpoint_signature="adaptive-contract",
            device=torch.device("cpu"),
        )

    resumed_calls = 0

    def resumed(*args, **kwargs):
        nonlocal resumed_calls
        resumed_calls += 1
        return 0.5

    monkeypatch.setattr(adaptive, "_train_one_epoch", resumed)
    result = adaptive.run_adaptive_training(
        ModelConfig(n_features=2, cnn_channels=4, lstm_hidden=4, dense_units=4),
        train_loader=[],
        val_loader=[],
        epochs=3,
        checkpoint_store=store,
        checkpoint_stage="adaptive",
        checkpoint_signature="adaptive-contract",
        device=torch.device("cpu"),
    )

    assert resumed_calls == 2
    assert result.checkpoint_resumed is True
    assert len(result.history) == 3
    assert set(result.q_values) == {0.0001, 0.0005, 0.001, 0.005}


def test_xgboost_continues_from_saved_boosting_round(tmp_path: Path):
    x = pd.DataFrame({"f0": np.linspace(0, 1, 80, dtype=np.float32)})
    y = pd.Series(2 * x["f0"] + 0.1)
    store = _store(tmp_path, model="xgboost")
    seed_model = XGBRegressor(
        n_estimators=3,
        max_depth=2,
        tree_method="hist",
        eval_metric="mae",
        n_jobs=1,
    )
    seed_model.fit(x.iloc[:60], y.iloc[:60], eval_set=[(x.iloc[60:], y.iloc[60:])], verbose=False)
    checkpoint_path = store.save_xgboost(
        seed_model.get_booster(),
        "final",
        signature="same-training-contract",
        completed_rounds=3,
        completed=False,
    )
    assert checkpoint_path is not None
    checkpoint_path.with_suffix(checkpoint_path.suffix + ".meta.json").unlink()

    result = fit_xgboost_resumable(
        {
            "n_estimators": 5,
            "max_depth": 2,
            "tree_method": "hist",
            "eval_metric": "mae",
            "n_jobs": 1,
        },
        x.iloc[:60],
        y.iloc[:60],
        x.iloc[60:],
        y.iloc[60:],
        store=store,
        stage="final",
        signature="same-training-contract",
    )

    assert result.resumed is True
    assert result.initial_rounds == 3
    assert result.completed_rounds == 5
    second = fit_xgboost_resumable(
        {
            "n_estimators": 5,
            "max_depth": 2,
            "tree_method": "hist",
            "eval_metric": "mae",
            "n_jobs": 1,
        },
        x.iloc[:60],
        y.iloc[:60],
        x.iloc[60:],
        y.iloc[60:],
        store=store,
        stage="final",
        signature="same-training-contract",
    )
    assert second.initial_rounds == second.completed_rounds == 5


def test_xgboost_early_stopping_state_survives_completed_resume(tmp_path: Path):
    train_x = pd.DataFrame({"f0": np.zeros(40, dtype=np.float32)})
    validation_x = pd.DataFrame({"f0": np.zeros(20, dtype=np.float32)})
    store = _store(tmp_path, model="xgboost")
    params = {
        "n_estimators": 20,
        "max_depth": 1,
        "learning_rate": 0.2,
        "tree_method": "hist",
        "eval_metric": "mae",
        "early_stopping_rounds": 2,
        "n_jobs": 1,
    }
    first = fit_xgboost_resumable(
        params,
        train_x,
        pd.Series(np.ones(40, dtype=np.float32)),
        validation_x,
        pd.Series(np.zeros(20, dtype=np.float32)),
        store=store,
        stage="early-stop",
        signature="early-stop-contract",
    )

    assert first.completed_rounds < 20
    assert first.model.get_booster().attr("solar_early_stopping_wait") == "2"
    second = fit_xgboost_resumable(
        params,
        train_x,
        pd.Series(np.ones(40, dtype=np.float32)),
        validation_x,
        pd.Series(np.zeros(20, dtype=np.float32)),
        store=store,
        stage="early-stop",
        signature="early-stop-contract",
    )
    assert second.resumed is True
    assert second.initial_rounds == second.completed_rounds == first.completed_rounds


def test_xgboost_restores_early_stopping_wait_after_interruption(tmp_path: Path):
    train_x = pd.DataFrame({"f0": np.zeros(40, dtype=np.float32)})
    validation_x = pd.DataFrame({"f0": np.zeros(20, dtype=np.float32)})
    train_y = pd.Series(np.ones(40, dtype=np.float32))
    validation_y = pd.Series(np.zeros(20, dtype=np.float32))
    seed = XGBRegressor(
        n_estimators=2,
        max_depth=1,
        learning_rate=0.2,
        tree_method="hist",
        eval_metric="mae",
        n_jobs=1,
    )
    seed.fit(train_x, train_y, eval_set=[(validation_x, validation_y)], verbose=False)
    seed.get_booster().set_attr(
        best_score="0.0",
        best_iteration="0",
        solar_early_stopping_best_score="0.0",
        solar_early_stopping_best_iteration="0",
        solar_early_stopping_wait="1",
    )
    store = _store(tmp_path, model="xgboost")
    store.save_xgboost(
        seed.get_booster(),
        "interrupted-early-stop",
        signature="interrupted-contract",
        completed_rounds=2,
        completed=False,
    )

    resumed = fit_xgboost_resumable(
        {
            "n_estimators": 20,
            "max_depth": 1,
            "learning_rate": 0.2,
            "tree_method": "hist",
            "eval_metric": "mae",
            "early_stopping_rounds": 2,
            "n_jobs": 1,
        },
        train_x,
        train_y,
        validation_x,
        validation_y,
        store=store,
        stage="interrupted-early-stop",
        signature="interrupted-contract",
    )

    assert resumed.initial_rounds == 2
    assert resumed.completed_rounds == 3
    assert resumed.model.get_booster().attr("solar_early_stopping_wait") == "2"
