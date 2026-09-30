"""전체 파이프라인에 CNN 학습·평가 결과를 전달하는 adapter."""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from solar_forecast.models.cnn_bilstm.evaluation import evaluate_and_analyze
from solar_forecast.models.cnn_bilstm.training_workflow import train_cnn_bilstm
from solar_forecast.pipeline.pipeline_config import PipelineConfig


class CnnTrainingAdapter:
    """Adapter exposing the CNN workflow to the application pipeline."""

    def execute(self, frame: pd.DataFrame, features: list[str], config: PipelineConfig, run_dir: Path) -> tuple[dict[str, Any], dict[str, Any]]:
        artifacts = train_cnn_bilstm(
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
            checkpoint_root=config.output_dir / ".checkpoints",
            optimizer_storage_path=config.output_dir / "optimization.db",
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
