"""Saved CNN replay preserves target units without training or refitting."""
from dataclasses import asdict

import numpy as np
import pandas as pd
import pytest
import torch

from solar_forecast.models.cnn_bilstm.network import CnnBiLstmNetworkConfig, build_cnn_bilstm_network
from solar_forecast.models.cnn_bilstm.sequence_data import SequenceConfig, prepare_dataset_splits
from solar_forecast.models.cnn_bilstm.training_workflow import _write_prediction_artifact
from tools.verify_benchmark_model_artifacts import _cnn_replay, _cnn_target_scale, _indexed, _load_observations


@pytest.mark.parametrize("transform", ["identity", "capacity_factor"])
def test_cnn_saved_replay_matches_exported_mwh_with_row_specific_capacity(tmp_path, monkeypatch, transform):
    import solar_forecast.models.cnn_bilstm.sequence_data as sequence_data

    frame = pd.DataFrame({
        "plant_id": "p", "plant": "P", "region": "R",
        "timestamp": pd.date_range("2024-01-01", periods=160, freq="h"),
        "f": np.arange(160, dtype=np.float32), "generation_mwh": np.float32(0.2),
        "rated_mw": np.linspace(0.5, 1.5, 160, dtype=np.float32),
        "energy_source": "solar", "quality_train_eligible": True,
    })
    cfg = SequenceConfig(sequence_length=4, batch_size=8, prediction_task="historical_forecast",
                         forecast_horizon_hours=1, target_transform=transform, capacity_column="rated_mw")
    splits = prepare_dataset_splits(frame, "generation_mwh", ["f"], cfg, "plant_id", "timestamp")
    network = CnnBiLstmNetworkConfig(n_features=splits.n_features, cnn_channels=4, kernel_size=3,
                                   lstm_hidden=4, lstm_layers=1, dense_units=4, dropout=0)
    torch.manual_seed(42)
    model = build_cnn_bilstm_network(network, device=torch.device("cpu"))
    checkpoint = {"config": asdict(network), "model_state": model.state_dict(), "feature_columns": ["f"],
                  "preprocessing": splits.train.preprocessing_state,
                  "target_transform": transform, "capacity_column": "rated_mw"}
    if transform == "identity":
        # Old MWh checkpoints did not carry target-transform fields.
        checkpoint.pop("target_transform")
        checkpoint.pop("capacity_column")
    model_path = tmp_path / "cnn.pt"
    torch.save(checkpoint, model_path)
    prediction_path = tmp_path / "test.csv"
    _write_prediction_artifact(model, splits.test, prediction_path, split="test", device=torch.device("cpu"))
    exported = _indexed(pd.read_csv(prediction_path)).iloc[::-1]
    values = {"feature_columns": ["f"], "target_column": "generation_mwh", "target_transform": transform,
              "capacity_column": "rated_mw", "forecast_horizon_hours": 1, "batch_size": 8,
              "energy_source_filter": "solar", "quality_filter_column": "quality_train_eligible"}
    source = tmp_path / "gold.csv"
    frame.to_csv(source, index=False)
    loaded = _load_observations(source, values, smoke=False)
    def forbid_refit(*args, **kwargs):
        raise AssertionError("replay must not fit preprocessing")
    monkeypatch.setattr(sequence_data, "fit_input_preprocessing", forbid_refit)
    predicted, evidence = _cnn_replay(loaded, values, {"checkpoint_path": str(model_path)}, exported, tmp_path)
    np.testing.assert_allclose(predicted, exported.y_pred, rtol=0, atol=1e-6)
    assert evidence["target_transform"] == transform


def test_cnn_replay_fails_closed_on_changed_or_missing_target_contract():
    values = {"target_transform": "capacity_factor", "capacity_column": "rated_mw"}
    checkpoint = dict(values)
    frame = pd.DataFrame({"rated_mw": [0.7, 0, -1, np.nan, np.inf]})
    scales = _cnn_target_scale(frame, values, checkpoint)
    assert scales[0] == np.float32(0.7)
    assert np.isnan(scales[1:]).all()
    for changed in ({}, {**checkpoint, "target_transform": "identity"}, {**checkpoint, "capacity_column": "other"}):
        with pytest.raises(ValueError):
            _cnn_target_scale(frame, values, changed)
    with pytest.raises(ValueError, match="requires rated_mw"):
        _cnn_target_scale(frame.drop(columns="rated_mw"), values, checkpoint)
    with pytest.raises(ValueError, match="Unsupported"):
        _cnn_target_scale(frame, {"target_transform": "unknown"}, {"target_transform": "unknown"})
