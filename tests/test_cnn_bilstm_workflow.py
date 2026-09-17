from pathlib import Path
import json
import re

import numpy as np
import optuna
import pandas as pd
import pytest
import torch

from solar_forecast.models.cnn_bilstm.network import CnnBiLstmNetworkConfig
from solar_forecast.models.cnn_bilstm.network import build_cnn_bilstm_network
from solar_forecast.models.cnn_bilstm.optimization import optimize_cnn_bilstm
from solar_forecast.models.cnn_bilstm.optimization import train_with_best_trial
from solar_forecast.models.cnn_bilstm.sequence_config import SequenceConfig
from solar_forecast.models.cnn_bilstm.sequence_data import LazyWindowSequenceDataset
from solar_forecast.models.cnn_bilstm.sequence_data import _EntitySeries
from solar_forecast.models.cnn_bilstm.sequence_data import _fit_and_transform_training_preprocessing
from solar_forecast.models.cnn_bilstm.sequence_data import prepare_dataset_splits
from solar_forecast.models.cnn_bilstm.sequence_data import prepare_datasets
from solar_forecast.models.cnn_bilstm.evaluation import compare_checkpoints
from solar_forecast.models.cnn_bilstm.evaluation import detect_outliers_from_predictions
from solar_forecast.models.cnn_bilstm.evaluation import evaluate_and_analyze
from solar_forecast.models.cnn_bilstm.training_workflow import train_cnn_bilstm
from solar_forecast.models.shared.optuna_study import OptimizationSettings


def _dummy_frame(n_rows: int = 120, n_features: int = 3) -> pd.DataFrame:
    rng = np.random.default_rng(0)
    features = rng.normal(size=(n_rows, n_features))
    target = features.sum(axis=1) + rng.normal(scale=0.1, size=n_rows)
    cols = {f"f{i}": features[:, i] for i in range(n_features)}
    cols["target"] = target
    return pd.DataFrame(cols)


def _short_seq_config() -> SequenceConfig:
    return SequenceConfig(sequence_length=5, test_size=0.2, val_size=0.2, batch_size=16, shuffle=True, num_workers=0)


def test_train_and_save_creates_timestamped_dir(tmp_path):
    frame = _dummy_frame()
    base_output = tmp_path / "checkpoints"
    cfg = _short_seq_config()
    artifacts = train_cnn_bilstm(
        frame,
        target_column="target",
        feature_columns=["f0", "f1", "f2"],
        sequence_config=cfg,
        output_dir=str(base_output),
        use_optuna=False,
        use_reinforcement=False,
        epochs=2,
    )

    run_dir = Path(artifacts["output_dir"])
    assert run_dir.parent == base_output
    assert re.match(r"\d{8}_\d{6}", run_dir.name)
    for fname in [
        "cnn_bilstm.pt",
        "metrics.json",
        "best_params.json",
        "validation_predictions.csv",
        "calibration_predictions.csv",
        "test_predictions.csv",
    ]:
        assert (run_dir / fname).exists()
    predictions = pd.read_csv(run_dir / "test_predictions.csv")
    assert set(
        [
            "timestamp",
            "plant_id",
            "region",
            "plant",
            "split",
            "y_true",
            "y_pred",
            "cnn_pred",
        ]
    ).issubset(predictions.columns)
    assert predictions["split"].eq("test").all()
    checkpoint = torch.load(run_dir / "cnn_bilstm.pt", weights_only=True)
    assert checkpoint["config"]["readout"] == "final_hidden"
    assert checkpoint["preprocessing"]["schema_version"] == 2
    assert checkpoint["preprocessing"]["target_transform"] == "identity"
    assert checkpoint["sequence_config"]["sequence_length"] == cfg.sequence_length
    assert json.loads((run_dir / "best_params.json").read_text())["readout"] == "final_hidden"


def test_compare_checkpoints_reads_nested_runs(tmp_path):
    frame = _dummy_frame()
    base_output = tmp_path / "nested"
    cfg = _short_seq_config()
    artifacts = train_cnn_bilstm(
        frame,
        target_column="target",
        feature_columns=["f0", "f1", "f2"],
        sequence_config=cfg,
        output_dir=str(base_output),
        use_optuna=False,
        use_reinforcement=False,
        epochs=2,
    )

    summary = compare_checkpoints(
        str(base_output),
        frame,
        target_column="target",
        feature_columns=["f0", "f1", "f2"],
        sequence_config=cfg,
    )
    assert not summary.empty
    assert summary["checkpoint"].str.endswith(".pt").all()
    assert Path(artifacts["checkpoint_path"]).exists()


def test_evaluate_and_analyze_saves_outputs(tmp_path):
    frame = _dummy_frame()
    base_output = tmp_path / "checkpoints"
    analysis_output = tmp_path / "analysis"
    cfg = _short_seq_config()
    artifacts = train_cnn_bilstm(
        frame,
        target_column="target",
        feature_columns=["f0", "f1", "f2"],
        sequence_config=cfg,
        output_dir=str(base_output),
        use_optuna=False,
        use_reinforcement=False,
        epochs=2,
    )

    analysis = evaluate_and_analyze(
        artifacts["checkpoint_path"],
        frame,
        target_column="target",
        feature_columns=["f0", "f1", "f2"],
        sequence_config=cfg,
        output_dir=str(analysis_output),
    )

    run_dir = Path(analysis["output_dir"])
    assert run_dir.parent == analysis_output
    assert (run_dir / "metrics.json").exists()
    assert (run_dir / "anomalies.csv").exists()
    anomalies = pd.read_csv(run_dir / "anomalies.csv")
    assert not anomalies.empty


def test_entity_sequences_never_cross_plants_and_split_chronologically():
    frame = pd.DataFrame(
        {
            "timestamp": list(pd.date_range("2025-01-01", periods=30, freq="h")) * 2,
            "plant_id": ["a"] * 30 + ["b"] * 30,
            "f0": list(range(30)) + list(range(100, 130)),
            "target": list(range(30)) + list(range(100, 130)),
        }
    )
    cfg = SequenceConfig(
        sequence_length=3,
        test_size=0.2,
        val_size=0.2,
        batch_size=128,
        shuffle=False,
        num_workers=0,
    )
    train, validation, test, _ = prepare_datasets(
        frame,
        "target",
        ["f0"],
        cfg,
        entity_column="plant_id",
        timestamp_column="timestamp",
    )
    for loader in (train, validation, test):
        for features, _ in loader:
            state = loader.preprocessing_state
            values = features.numpy()[:, :, 0] * state["feature_scales"]["f0"] + state["feature_means"]["f0"]
            assert all((row < 50).all() or (row > 50).all() for row in values)
    train_targets = next(iter(train))[1].numpy()
    validation_targets = next(iter(validation))[1].numpy()
    test_targets = next(iter(test))[1].numpy()
    for lower, upper in ((0, 50), (100, 150)):
        train_group = train_targets[(train_targets > lower) & (train_targets < upper)]
        validation_group = validation_targets[(validation_targets > lower) & (validation_targets < upper)]
        test_group = test_targets[(test_targets > lower) & (test_targets < upper)]
        assert train_group.max() < validation_group.min() < test_group.min()


def test_anomaly_threshold_is_frozen_from_calibration_not_test_ranking():
    calibration = np.linspace(-1.0, 1.0, 100)
    result = detect_outliers_from_predictions(
        np.array([0.0, 0.0, 0.0]),
        np.array([0.1, 1.5, 2.0]),
        contamination=0.05,
        calibration_residuals=calibration,
    )
    assert result["anomaly_threshold"].nunique() == 1
    assert result["threshold_source"].eq(
        "frozen_independent_calibration_absolute_residual_quantile"
    ).all()
    assert result["is_outlier"].tolist() == [False, True, True]


def test_lazy_windows_keep_all_missing_train_feature_as_zero_plus_mask():
    frame = _dummy_frame(100, 1)
    frame["never_observed"] = np.nan
    cfg = SequenceConfig(
        sequence_length=5,
        batch_size=16,
        shuffle=False,
        append_missing_indicators=True,
    )
    splits = prepare_dataset_splits(
        frame,
        "target",
        ["f0", "never_observed"],
        cfg,
    )
    assert isinstance(splits.train.dataset, LazyWindowSequenceDataset)
    assert splits.n_features == 4
    features, _ = next(iter(splits.train))
    assert features[:, :, 1].eq(0).all()
    assert features[:, :, 3].eq(1).all()
    state = splits.train.preprocessing_state
    assert state["all_missing_training_features"] == ["never_observed"]
    assert state["temporal_split"]["window_materialization"] == "lazy_per_batch"


@pytest.mark.parametrize("read_only", [True, False])
@pytest.mark.parametrize("append_missing_indicators", [True, False])
def test_imputation_owns_buffer_and_preserves_input(read_only, append_missing_indicators):
    source = np.array([[1, np.nan], [np.nan, 10], [5, 30], [1000, 9000]], dtype=np.float32)
    original = source.copy()
    source.setflags(write=not read_only)
    series = _EntitySeries(
        features=source,
        targets=np.zeros(4, dtype=np.float32),
        target_positions={},
        train_feature_rows=np.array([True, True, True, False]),
        plant_id="test-plant", region="test-region", plant="test-plant",
        timestamps=np.arange(4),
    )

    state = _fit_and_transform_training_preprocessing(
        [series], ["first", "second"], append_missing_indicators=append_missing_indicators,
    )

    assert state["feature_medians"] == {"first": 3.0, "second": 20.0}
    reconstructed = series.features[:, :2] * np.array(list(state["feature_scales"].values())) + np.array(list(state["feature_means"].values()))
    np.testing.assert_allclose(reconstructed, [[1, 20], [3, 10], [5, 30], [1000, 9000]], rtol=1e-6)
    np.testing.assert_array_equal(source, original)
    assert source.flags.writeable == (not read_only)
    assert series.features.flags.writeable
    assert not np.shares_memory(source, series.features)
    if append_missing_indicators:
        np.testing.assert_array_equal(series.features[:, 2:], [[0, 1], [1, 0], [0, 0], [0, 0]])


def test_historical_cnn_context_matches_tabular_forecast_and_excludes_future_inputs():
    from solar_forecast.evaluation.forecast_samples import build_forecast_samples

    frame = pd.DataFrame({
        "timestamp": pd.date_range("2025-01-01", periods=300, freq="h"),
        "plant_id": "plant-a",
        "weather": np.arange(300, dtype=float),
        "target": np.arange(300, dtype=float) + 1000,
    }).drop(index=250)
    cfg = SequenceConfig(
        sequence_length=12, batch_size=128, shuffle=False,
        append_missing_indicators=False,
        prediction_task="historical_forecast", forecast_horizon_hours=6,
    )
    splits = prepare_dataset_splits(frame, "target", ["weather"], cfg, "plant_id", "timestamp")
    tabular = build_forecast_samples(frame, ["weather"], "target", 6).set_index("timestamp")
    for dataset in (splits.train.dataset, splits.validation.dataset, splits.calibration.dataset, splits.test.dataset):
        context = dataset.context_frame(0, len(dataset))
        for index, row in context.iterrows():
            features, target = dataset[index]
            origin = pd.Timestamp(row["forecast_origin"])
            observed = tabular.loc[pd.Timestamp(row["timestamp"])]
            assert origin == observed["forecast_origin"]
            state = splits.train.preprocessing_state
            reconstructed = features[:, 0].numpy() * state["feature_scales"]["weather"] + state["feature_means"]["weather"]
            assert reconstructed[-1] == pytest.approx(observed["weather"], abs=1e-4)
            assert float(target) == observed["target"]
            assert row["persistence_pred"] == observed["persistence_pred"]
            np.testing.assert_allclose(np.diff(reconstructed), np.ones(11), atol=1e-4)
    training = splits.train.dataset.series[0]
    latest_train_origin = training.origin_positions[training.target_positions["train"]].max()
    assert not training.train_feature_rows[latest_train_origin + 1:].any()
    assert splits.split_metadata["boundaries"]["gap_hours"] == 6


def test_historical_lookbacks_keep_common_split_calendar():
    frame = pd.DataFrame({
        "timestamp": pd.date_range("2025-01-01", periods=300, freq="h"),
        "plant_id": "plant-a", "weather": np.arange(300), "target": np.arange(300),
    })
    metadata = []
    for length in (6, 48):
        splits = prepare_dataset_splits(
            frame, "target", ["weather"],
            SequenceConfig(sequence_length=length, prediction_task="historical_forecast", forecast_horizon_hours=12),
            "plant_id", "timestamp",
        )
        metadata.append(splits.split_metadata)
    assert metadata[0]["boundaries"] == metadata[1]["boundaries"]
    assert metadata[0]["test_period"] == metadata[1]["test_period"]


@pytest.mark.parametrize("layers", [1, 2])
def test_final_hidden_readout_uses_both_top_layer_final_states(layers):
    torch.manual_seed(17)
    config = CnnBiLstmNetworkConfig(
        n_features=3, cnn_channels=4, lstm_hidden=5,
        lstm_layers=layers, dense_units=4, dropout=0.0, readout="final_hidden",
    )
    model = build_cnn_bilstm_network(config).eval()
    # Expose the sequence summary directly, independently of dense head weights.
    model.head = torch.nn.Identity()
    inputs = torch.randn(2, 7, 3)
    with torch.no_grad():
        convolved = model.conv(inputs.transpose(1, 2)).transpose(1, 2)
        output, (hidden, _) = model.lstm(convolved)
        actual = model(inputs)
    torch.testing.assert_close(actual[:, :5], hidden[-2], rtol=0, atol=0)
    torch.testing.assert_close(actual[:, 5:], hidden[-1], rtol=0, atol=0)
    torch.testing.assert_close(actual[:, 5:], output[:, 0, 5:], rtol=0, atol=0)
    assert not torch.allclose(actual[:, 5:], output[:, -1, 5:])


def test_omitted_readout_retains_legacy_last_output_semantics():
    torch.manual_seed(19)
    historical_config = {
        "n_features": 3, "cnn_channels": 4, "kernel_size": 3,
        "lstm_hidden": 5, "lstm_layers": 2, "dense_units": 4, "dropout": 0.0,
    }
    model = build_cnn_bilstm_network(CnnBiLstmNetworkConfig(**historical_config)).eval()
    inputs = torch.randn(2, 7, 3)
    with torch.no_grad():
        convolved = model.conv(inputs.transpose(1, 2)).transpose(1, 2)
        output, _ = model.lstm(convolved)
        expected = model.head(output[:, -1]).squeeze(-1)
        actual = model(inputs)
    assert model.readout == "last_output"
    torch.testing.assert_close(actual, expected, rtol=0, atol=0)


@pytest.mark.parametrize("readout", ["last_output", "final_hidden"])
def test_checkpoint_preserves_readout_and_exact_predictions(tmp_path, readout):
    torch.manual_seed(23)
    config = CnnBiLstmNetworkConfig(
        n_features=3, cnn_channels=4, lstm_hidden=5,
        lstm_layers=2, dense_units=4, dropout=0.0, readout=readout,
    )
    model = build_cnn_bilstm_network(config).eval()
    inputs = torch.randn(2, 7, 3)
    path = tmp_path / "model.pt"
    torch.save({"config": config.__dict__, "model_state": model.state_dict()}, path)
    checkpoint = torch.load(path, weights_only=True)
    restored = build_cnn_bilstm_network(CnnBiLstmNetworkConfig(**checkpoint["config"])).eval()
    restored.load_state_dict(checkpoint["model_state"])
    with torch.no_grad():
        torch.testing.assert_close(restored(inputs), model(inputs), rtol=0, atol=0)
    assert restored.readout == readout


def test_network_rejects_unknown_readout():
    with pytest.raises(ValueError, match="readout"):
        CnnBiLstmNetworkConfig(n_features=3, readout="mean")


def _tiny_network_search_space(readout=None):
    values = {
        "cnn_channels": 4, "kernel_size": 3, "lstm_hidden": 4,
        "lstm_layers": 1, "dense_units": 4, "dropout": 0.0,
        "lr": 0.001, "weight_decay": 0.0,
    }
    space = {name: {"type": "fixed", "value": value} for name, value in values.items()}
    if readout is not None:
        space["readout"] = {"type": "categorical", "choices": [readout]}
    return space


@pytest.mark.parametrize("readout", [None, "last_output", "final_hidden"])
def test_trial_and_final_fit_use_the_same_selected_readout(readout):
    result = train_with_best_trial(
        _dummy_frame(), "target", feature_columns=["f0", "f1", "f2"],
        sequence_config=_short_seq_config(), n_trials=1, trial_epochs=1,
        epochs=1, device=torch.device("cpu"),
        optimizer_parameter_space=_tiny_network_search_space(readout),
    )
    expected = readout or "final_hidden"
    assert result["study"].best_params["readout"] == expected
    assert result["best_params"]["readout"] == expected
    assert result["model_config"].readout == expected
    assert result["model"].readout == expected


def test_readout_study_does_not_reuse_legacy_or_different_readout_trials(tmp_path):
    storage_path = tmp_path / "optimizer.db"
    settings = OptimizationSettings(
        enabled=True, study_name="cnn_legacy", storage_path=storage_path,
        max_trials=1, timeout_seconds=60, seed=42, startup_trials=0,
        pruner_startup_trials=0, pruner_warmup_steps=0,
    )
    legacy = optuna.create_study(
        study_name=settings.study_name, storage=f"sqlite:///{storage_path}",
    )
    legacy.add_trial(optuna.trial.create_trial(value=-123.0))
    common = {
        "feature_columns": ["f0", "f1", "f2"],
        "sequence_config": _short_seq_config(), "trial_epochs": 1,
        "settings": settings, "device": torch.device("cpu"),
    }
    final_hidden = optimize_cnn_bilstm(
        _dummy_frame(), "target", artifact_dir=tmp_path / "final_hidden",
        optimizer_parameter_space=_tiny_network_search_space(), **common,
    )
    last_output = optimize_cnn_bilstm(
        _dummy_frame(), "target", artifact_dir=tmp_path / "last_output",
        optimizer_parameter_space=_tiny_network_search_space("last_output"), **common,
    )
    assert len({legacy.study_name, final_hidden.study_name, last_output.study_name}) == 3
    assert legacy.best_value == -123.0
    assert final_hidden.best_params["readout"] == "final_hidden"
    assert last_output.best_params["readout"] == "last_output"


def test_checkpoint_evaluation_reuses_saved_statistics_without_refit(tmp_path, monkeypatch):
    from solar_forecast.models.cnn_bilstm.evaluation import _checkpoint_loaders
    import solar_forecast.models.cnn_bilstm.sequence_data as sequence_data

    frame = _dummy_frame(160)
    cfg = _short_seq_config()
    splits = prepare_dataset_splits(frame, "target", ["f0", "f1", "f2"], cfg)
    data = {
        "config": {"n_features": splits.n_features},
        "preprocessing": splits.train.preprocessing_state,
        "feature_columns": ["f0", "f1", "f2"],
        "sequence_config": cfg.__dict__,
        "target_column": "target",
    }
    def no_fit(*args, **kwargs):
        raise AssertionError("Checkpoint replay must never fit preprocessing")
    monkeypatch.setattr(sequence_data, "fit_input_preprocessing", no_fit)
    changed = frame.copy()
    changed.loc[:20, "f0"] = 1e6
    replay = _checkpoint_loaders(data, changed, "target", None, None)
    assert replay.train.preprocessing_state["feature_means"] == splits.train.preprocessing_state["feature_means"]
    np.testing.assert_array_equal(replay.test.dataset[0][0], splits.test.dataset[0][0])
    with pytest.raises(ValueError, match="feature order"):
        _checkpoint_loaders(data, changed, "target", ["f1", "f0", "f2"], None)
    with pytest.raises(ValueError, match="sequence config"):
        _checkpoint_loaders(data, changed, "target", None, SequenceConfig(sequence_length=12))


def test_frozen_calendar_checkpoint_keeps_boundaries_when_data_grows():
    frame = pd.DataFrame({
        "timestamp": pd.date_range("2024-01-01", periods=360, freq="h"),
        "plant_id": "a", "feature": np.arange(360), "target": np.arange(360),
    })
    cfg = SequenceConfig(
        sequence_length=6, prediction_task="historical_forecast", forecast_horizon_hours=1,
        train_end="2024-01-05T23:00:00", validation_end="2024-01-08T23:00:00",
        calibration_end="2024-01-10T23:00:00", test_end="2024-01-12T23:00:00",
    )
    splits = prepare_dataset_splits(frame.iloc[:300], "target", ["feature"], cfg, "plant_id", "timestamp")
    replay = prepare_dataset_splits(
        frame, "target", ["feature"], cfg, "plant_id", "timestamp",
        preprocessing_state=splits.train.preprocessing_state,
    )
    assert replay.split_metadata["boundaries"] == splits.split_metadata["boundaries"]
    assert replay.split_metadata["counts"] == splits.split_metadata["counts"]
    assert replay.split_metadata["test_period"]["end"] == "2024-01-12T23:00:00"
