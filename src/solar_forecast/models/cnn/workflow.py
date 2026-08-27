"""High-level workflows for training, evaluation, and anomaly analysis."""
from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence

import numpy as np
import pandas as pd
import torch
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score

from .data import SequenceConfig, prepare_dataset_splits, prepare_datasets
from .model import ModelConfig, build_model
from .optimization import _evaluate, _train_one_epoch, save_study_results, train_with_best_trial
from .adaptive import BanditConfig, run_adaptive_training


def _timestamped_dir(base_dir: Path) -> Path:
    base_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = base_dir / stamp
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def train_and_save(
    frame: pd.DataFrame,
    target_column: str,
    feature_columns: Optional[Sequence[str]] = None,
    sequence_config: Optional[SequenceConfig] = None,
    n_trials: int = 10,
    output_dir: str = "cnn_bilstm/output/checkpoints",
    use_optuna: bool = True,
    use_reinforcement: bool = False,
    epochs: int = 50,
    entity_column: Optional[str] = None,
    timestamp_column: Optional[str] = None,
) -> Dict[str, object]:
    """Train the model with Optuna and/or reinforcement learning then persist artifacts.

    All artifacts for a run are stored in a timestamped subdirectory within ``output_dir``.
    """

    cfg = sequence_config or SequenceConfig()
    run_dir = _timestamped_dir(Path(output_dir))

    if use_optuna:
        result = train_with_best_trial(
            frame,
            target_column,
            feature_columns=feature_columns,
            sequence_config=cfg,
            n_trials=n_trials,
            entity_column=entity_column,
            timestamp_column=timestamp_column,
        )
        model, model_cfg = result["model"], result["model_config"]
        study = result["study"]
        save_study_results(study, str(run_dir / "optuna_best.json"))
    else:
        # Fallback to deterministic config
        loaders = prepare_dataset_splits(
            frame, target_column, feature_columns, cfg, entity_column, timestamp_column
        )
        train_loader, val_loader, test_loader = loaders.train, loaders.validation, loaders.test
        model_cfg = ModelConfig(n_features=loaders.n_features)
        device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
        model = build_model(model_cfg, device=device)
        criterion = torch.nn.MSELoss()
        optimizer = torch.optim.Adam(model.parameters(), lr=1e-3)
        for _ in range(epochs):
            _train_one_epoch(model, train_loader, criterion, optimizer, device)
        metrics = _evaluate(model, test_loader, criterion, device)
        result = {"model": model, "model_config": model_cfg, "metrics": metrics, "best_params": {}}

    preprocessing_state = (
        result.get("preprocessing")
        if use_optuna
        else getattr(train_loader, "preprocessing_state", None)
    )
    temporal_split = (
        preprocessing_state.get("temporal_split")
        if isinstance(preprocessing_state, dict)
        else None
    )

    if use_reinforcement:
        train_loader, val_loader, _, _ = prepare_datasets(
            frame, target_column, feature_columns, cfg, entity_column, timestamp_column
        )
        reinforcement = run_adaptive_training(
            model_cfg, train_loader, val_loader, epochs=30, bandit_cfg=BanditConfig(actions=[1e-4, 5e-4, 1e-3, 5e-3])
        )
        model = reinforcement.model
        with open(run_dir / "reinforcement_history.json", "w", encoding="utf-8") as f:
            json.dump(reinforcement.history, f, indent=2)
        with open(run_dir / "bandit_q_values.json", "w", encoding="utf-8") as f:
            json.dump(reinforcement.q_values, f, indent=2)

    checkpoint_path = run_dir / "cnn_bilstm.pt"
    torch.save(
        {
            "model_state": model.state_dict(),
            "config": model_cfg.__dict__,
            "feature_columns": list(feature_columns or []),
            "preprocessing": preprocessing_state,
        },
        checkpoint_path,
    )
    with open(run_dir / "metrics.json", "w", encoding="utf-8") as f:
        json.dump(result.get("metrics", {}), f, indent=2)
    with open(run_dir / "best_params.json", "w", encoding="utf-8") as f:
        json.dump(result.get("best_params", {}), f, indent=2)

    return {
        "model": model,
        "config": model_cfg,
        "metrics": result.get("metrics", {}),
        "output_dir": str(run_dir),
        "checkpoint_path": str(checkpoint_path),
        "temporal_split": temporal_split,
    }


def load_checkpoint(path: str, device: Optional[torch.device] = None):
    data = torch.load(path, map_location=device or "cpu")
    cfg = ModelConfig(**data["config"])
    model = build_model(cfg, device=device or torch.device("cpu"))
    model.load_state_dict(data["model_state"])
    model.eval()
    return model, cfg


def evaluate_model(
    model: torch.nn.Module,
    data_loader,
    device: Optional[torch.device] = None,
) -> Dict[str, float]:
    device = device or torch.device("cuda" if torch.cuda.is_available() else "cpu")
    criterion = torch.nn.MSELoss()
    metrics = _evaluate(model.to(device), data_loader, criterion, device)
    return metrics


def compare_checkpoints(
    checkpoint_dir: str,
    frame: pd.DataFrame,
    target_column: str,
    feature_columns: Optional[Sequence[str]] = None,
    sequence_config: Optional[SequenceConfig] = None,
) -> pd.DataFrame:
    """Load all checkpoints in a directory and compare their metrics."""

    cfg = sequence_config or SequenceConfig()
    _, _, test_loader, _ = prepare_datasets(frame, target_column, feature_columns, cfg)
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")

    rows: List[Dict[str, object]] = []
    for checkpoint in Path(checkpoint_dir).rglob("*.pt"):
        model, model_cfg = load_checkpoint(str(checkpoint), device=device)
        metrics = evaluate_model(model, test_loader, device=device)
        rows.append({"checkpoint": checkpoint.name, **metrics, **model_cfg.__dict__})

    return pd.DataFrame(rows).sort_values(by="loss")


def detect_outliers_from_predictions(
    y_true: np.ndarray,
    y_pred: np.ndarray,
    contamination: float = 0.05,
    *,
    calibration_residuals: np.ndarray | None = None,
) -> pd.DataFrame:
    """Apply a frozen calibration threshold; never rank the evaluated set itself."""

    if not 0 < contamination < 1:
        raise ValueError("contamination must be between zero and one")
    if calibration_residuals is None:
        raise ValueError("Independent calibration residuals are required")
    residuals = y_true - y_pred
    calibration = np.asarray(calibration_residuals, dtype=float)
    calibration = np.abs(calibration[np.isfinite(calibration)])
    if len(calibration) < 5:
        raise ValueError("At least five calibration residuals are required for anomaly thresholding")
    quantile = min(1.0, np.ceil((len(calibration) + 1) * (1 - contamination)) / len(calibration))
    threshold = float(np.quantile(calibration, quantile, method="higher"))
    return pd.DataFrame(
        {
            "y_true": y_true,
            "y_pred": y_pred,
            "residual": residuals,
            "absolute_residual": np.abs(residuals),
            "anomaly_threshold": threshold,
            "threshold_source": "frozen_independent_calibration_absolute_residual_quantile",
            "is_outlier": np.abs(residuals) > threshold,
        }
    )


def evaluate_and_analyze(
    checkpoint_path: str,
    frame: pd.DataFrame,
    target_column: str,
    feature_columns: Optional[Sequence[str]] = None,
    sequence_config: Optional[SequenceConfig] = None,
    contamination: float = 0.05,
    output_dir: Optional[str] = None,
    entity_column: Optional[str] = None,
    timestamp_column: Optional[str] = None,
) -> Dict[str, object]:
    """Load a checkpoint, compute metrics, and perform outlier analysis.

    If ``output_dir`` is provided, anomalies and metrics are saved in a timestamped subdirectory.
    """

    cfg = sequence_config or SequenceConfig()
    loaders = prepare_dataset_splits(
        frame,
        target_column,
        feature_columns,
        cfg,
        entity_column,
        timestamp_column,
    )
    calibration_true_all: List[np.ndarray] = []
    calibration_pred_all: List[np.ndarray] = []
    y_true_all: List[np.ndarray] = []
    y_pred_all: List[np.ndarray] = []

    model, _ = load_checkpoint(checkpoint_path)
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    model = model.to(device)
    model.eval()
    with torch.no_grad():
        for X, y in loaders.calibration:
            X = X.to(device)
            preds = model(X).cpu().numpy()
            calibration_true_all.append(y.numpy())
            calibration_pred_all.append(preds)
        for X, y in loaders.test:
            X = X.to(device)
            preds = model(X).cpu().numpy()
            y_true_all.append(y.numpy())
            y_pred_all.append(preds)

    y_true = np.concatenate(y_true_all)
    y_pred = np.concatenate(y_pred_all)
    calibration_residuals = np.concatenate(calibration_true_all) - np.concatenate(calibration_pred_all)
    mse = mean_squared_error(y_true, y_pred)
    metrics = {
        "mae": mean_absolute_error(y_true, y_pred),
        "rmse": float(np.sqrt(mse)),
        "r2": r2_score(y_true, y_pred),
    }
    anomaly_df = detect_outliers_from_predictions(
        y_true,
        y_pred,
        contamination=contamination,
        calibration_residuals=calibration_residuals,
    )
    result: Dict[str, object] = {
        "metrics": metrics,
        "anomalies": anomaly_df,
        "temporal_split": loaders.split_metadata,
    }

    if output_dir:
        run_dir = _timestamped_dir(Path(output_dir))
        with open(run_dir / "metrics.json", "w", encoding="utf-8") as f:
            json.dump(metrics, f, indent=2)
        anomaly_df.to_csv(run_dir / "anomalies.csv", index=False)
        result["output_dir"] = str(run_dir)

    return result


def save_dataframe(df: pd.DataFrame, path: str) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False)
