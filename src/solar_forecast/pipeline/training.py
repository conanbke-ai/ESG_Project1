from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from solar_forecast.models.cnn.workflow import evaluate_and_analyze, train_and_save
from .config import PipelineConfig


class CnnTrainingAdapter:
    """Adapter exposing the CNN workflow to the application pipeline."""

    def execute(self, frame: pd.DataFrame, features: list[str], config: PipelineConfig, run_dir: Path) -> tuple[dict[str, Any], dict[str, Any]]:
        artifacts = train_and_save(
            frame,
            target_column=config.target_column,
            feature_columns=features,
            sequence_config=config.sequence,
            n_trials=config.n_trials,
            output_dir=str(run_dir / "model"),
            use_optuna=config.use_optuna,
            optimizer_timeout_seconds=config.optimizer_timeout_seconds,
            use_reinforcement=config.use_reinforcement,
            epochs=config.epochs,
        )
        analysis = evaluate_and_analyze(
            str(artifacts["checkpoint_path"]),
            frame,
            target_column=config.target_column,
            feature_columns=features,
            sequence_config=config.sequence,
            contamination=config.contamination,
            output_dir=None,
        )
        return artifacts, analysis


def train_and_evaluate(frame: pd.DataFrame, features: list[str], config: PipelineConfig, run_dir: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    return CnnTrainingAdapter().execute(frame, features, config, run_dir)
