from pathlib import Path

import numpy as np
import pandas as pd
import torch

from solar_forecast.models.cnn.config import SequenceConfig
from solar_forecast.models.cnn.workflow import train_and_save
from solar_forecast.models.optimization import (
    OptimizationSettings,
    OptunaStudyService,
)
from solar_forecast.models.xgboost_optimization import (
    XGBoostHyperparameterOptimizer,
)


def _settings(tmp_path: Path, study_name: str, *, max_trials: int = 1):
    return OptimizationSettings(
        enabled=True,
        study_name=study_name,
        storage_path=tmp_path / f"{study_name}.db",
        max_trials=max_trials,
        timeout_seconds=30,
        seed=42,
        startup_trials=0,
        pruner_startup_trials=0,
        pruner_warmup_steps=0,
    )


def test_optuna_study_resumes_without_exceeding_max_total_trials(tmp_path: Path):
    settings = _settings(tmp_path, "resume_test", max_trials=2)

    def objective(trial):
        value = trial.suggest_float("value", -1.0, 1.0)
        return (value - 0.25) ** 2

    first = OptunaStudyService(settings, project_root=tmp_path).run(
        objective, tmp_path / "first"
    )
    second = OptunaStudyService(settings, project_root=tmp_path).run(
        objective, tmp_path / "second"
    )

    assert first.existing_trials == 0
    assert first.executed_trials == 2
    assert second.existing_trials == 2
    assert second.executed_trials == 0
    assert second.summary_path.exists()
    assert second.trials_path.exists()


def test_xgboost_optimizer_uses_validation_and_persists_artifacts(tmp_path: Path):
    x = np.linspace(0.0, 1.0, 100, dtype=np.float32)
    frame = pd.DataFrame({"f0": x, "target": 2 * x + 0.1})
    optimizer = XGBoostHyperparameterOptimizer(
        _settings(tmp_path, "xgb_test"),
        {
            "seed": 42,
            "n_jobs": 1,
            "optimizer": {
                "trial_max_estimators": 5,
                "early_stopping_rounds": 2,
                "tuning_train_max_rows": 80,
                "tuning_validation_max_rows": 20,
            },
        },
    )
    result = optimizer.optimize(
        frame.iloc[:80],
        frame.iloc[80:],
        feature_columns=["f0"],
        target_column="target",
        artifact_dir=tmp_path / "xgb_artifacts",
    )

    assert result.tuning_train_rows == 80
    assert result.tuning_validation_rows == 20
    assert result.run.study.best_value >= 0
    assert result.run.summary_path.exists()


def test_cnn_optimizer_handles_missing_mask_dimension_and_saves_study(tmp_path: Path):
    rng = np.random.default_rng(7)
    frame = pd.DataFrame(
        {
            "f0": rng.normal(size=120),
            "all_missing": np.nan,
            "target": rng.normal(size=120),
        }
    )
    artifacts = train_and_save(
        frame,
        target_column="target",
        feature_columns=["f0", "all_missing"],
        sequence_config=SequenceConfig(
            sequence_length=5,
            val_size=0.2,
            calibration_size=0.1,
            test_size=0.2,
            batch_size=16,
            shuffle=False,
            append_missing_indicators=True,
        ),
        n_trials=1,
        output_dir=str(tmp_path / "cnn"),
        use_optuna=True,
        epochs=1,
        optimizer_settings=_settings(tmp_path, "cnn_test"),
        optimizer_trial_epochs=1,
        early_stopping_patience=1,
        optimizer_max_train_sequences=20,
        optimizer_max_validation_sequences=10,
    )

    checkpoint = torch.load(artifacts["checkpoint_path"], map_location="cpu")
    assert checkpoint["config"]["n_features"] == 4
    run_dir = Path(artifacts["output_dir"])
    assert (run_dir / "optimization_summary.json").exists()
    assert (run_dir / "optimization_trials.csv").exists()
