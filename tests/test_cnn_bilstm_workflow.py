from pathlib import Path
import re

import numpy as np
import pandas as pd
import pytest

from solar_forecast.models.cnn_bilstm.sequence_config import SequenceConfig
from solar_forecast.models.cnn_bilstm.sequence_data import LazyWindowSequenceDataset
from solar_forecast.models.cnn_bilstm.sequence_data import _EntitySeries
from solar_forecast.models.cnn_bilstm.sequence_data import _fit_and_transform_training_medians
from solar_forecast.models.cnn_bilstm.sequence_data import prepare_dataset_splits
from solar_forecast.models.cnn_bilstm.sequence_data import prepare_datasets
from solar_forecast.models.cnn_bilstm.evaluation import compare_checkpoints
from solar_forecast.models.cnn_bilstm.evaluation import detect_outliers_from_predictions
from solar_forecast.models.cnn_bilstm.evaluation import evaluate_and_analyze
from solar_forecast.models.cnn_bilstm.training_workflow import train_cnn_bilstm


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
            values = features.numpy()[:, :, 0]
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

    state = _fit_and_transform_training_medians(
        [series], ["first", "second"], append_missing_indicators=append_missing_indicators,
    )

    assert state["feature_medians"] == {"first": 3.0, "second": 20.0}
    np.testing.assert_array_equal(series.features[:, :2], [[1, 20], [3, 10], [5, 30], [1000, 9000]])
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
            assert float(features[-1, 0]) == observed["weather"]
            assert float(target) == observed["target"]
            assert row["persistence_pred"] == observed["persistence_pred"]
            np.testing.assert_array_equal(np.diff(features[:, 0]), np.ones(11))
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
