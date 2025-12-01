from pathlib import Path
import re

import numpy as np
import pandas as pd

from cnn_bilstm.data_utils import SequenceConfig
from cnn_bilstm.workflows import compare_checkpoints, evaluate_and_analyze, train_and_save


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
    artifacts = train_and_save(
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
    for fname in ["cnn_bilstm.pt", "metrics.json", "best_params.json"]:
        assert (run_dir / fname).exists()


def test_compare_checkpoints_reads_nested_runs(tmp_path):
    frame = _dummy_frame()
    base_output = tmp_path / "nested"
    cfg = _short_seq_config()
    artifacts = train_and_save(
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
    artifacts = train_and_save(
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
