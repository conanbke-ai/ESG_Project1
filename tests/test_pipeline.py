from pathlib import Path

import numpy as np
import pandas as pd

from solar_forecast.models.cnn.config import SequenceConfig
from solar_forecast.pipeline import PipelineConfig, run_pipeline
from solar_forecast.pipeline.dataset import discover_latest_file
from solar_forecast.pipeline.preprocessing import TrainFittedMedianImputer, preprocess_dataset


def test_preprocessing_selects_numeric_and_cleans_rows():
    frame = pd.DataFrame({"feature": [1, None, 3], "label": ["a", "b", "c"], "target": [2, 4, None]})
    result = preprocess_dataset(frame, "target")
    assert result.feature_columns == ["feature"]
    assert len(result.frame) == 2
    assert not result.frame.isna().any().any()


def test_train_fitted_imputer_does_not_use_future_values():
    train = pd.DataFrame({"feature": [1.0, np.nan, 3.0]})
    future = pd.DataFrame({"feature": [1_000_000.0, np.nan]})
    imputer = TrainFittedMedianImputer().fit(train, ["feature"])
    transformed = imputer.transform(future, ["feature"])
    assert transformed.iloc[1]["feature"] == 2.0


def test_latest_file_discovery(tmp_path):
    old = tmp_path / "old.csv"
    new = tmp_path / "new.csv"
    old.write_text("x\n1", encoding="utf-8")
    new.write_text("x\n2", encoding="utf-8")
    old.touch()
    new.touch()
    assert discover_latest_file(tmp_path) in {old, new}


def test_full_pipeline_creates_report(tmp_path):
    rng = np.random.default_rng(7)
    frame = pd.DataFrame({"f1": rng.normal(size=80), "f2": rng.normal(size=80)})
    frame["target"] = frame.f1 + frame.f2
    source = tmp_path / "input.csv"
    frame.to_csv(source, index=False)
    result = run_pipeline(PipelineConfig(
        target_column="target", data_path=source, output_dir=tmp_path / "runs",
        sequence=SequenceConfig(sequence_length=5, batch_size=16), use_optuna=False, epochs=1,
        artifact_level="debug",
    ))
    assert result.report_path.exists()
    assert result.processed_path.exists()
    assert result.checkpoint_path.exists()
    assert (result.run_dir / "manifest.json").exists()
