"""CNN-BiLSTM utilities with Optuna tuning and reinforcement learning."""

from .model import CNNBiLSTM, ModelConfig, build_model
from .data_utils import SequenceConfig, prepare_datasets, sequence_from_csv
from .optuna_search import run_study, train_with_best_trial, save_study_results
from .workflows import (
    compare_checkpoints,
    detect_outliers_from_predictions,
    evaluate_and_analyze,
    evaluate_model,
    load_checkpoint,
    save_dataframe,
    train_and_save,
)
from .reinforcement import BanditConfig, run_adaptive_training

__all__ = [
    "CNNBiLSTM",
    "ModelConfig",
    "build_model",
    "SequenceConfig",
    "prepare_datasets",
    "sequence_from_csv",
    "run_study",
    "train_with_best_trial",
    "save_study_results",
    "compare_checkpoints",
    "detect_outliers_from_predictions",
    "evaluate_and_analyze",
    "evaluate_model",
    "load_checkpoint",
    "save_dataframe",
    "train_and_save",
    "BanditConfig",
    "run_adaptive_training",
]
