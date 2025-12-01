"""Optuna study utilities for the CNN-BiLSTM model."""
from __future__ import annotations

import json
from typing import Dict, Optional

import numpy as np
import optuna
import torch
import torch.nn as nn
from optuna.trial import Trial
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from torch.utils.data import DataLoader

from .data_utils import SequenceConfig, prepare_datasets
from .model import CNNBiLSTM, ModelConfig, build_model


def _train_one_epoch(
    model: CNNBiLSTM,
    loader: DataLoader,
    criterion: nn.Module,
    optimizer: torch.optim.Optimizer,
    device: torch.device,
) -> float:
    model.train()
    running_loss = 0.0
    for X, y in loader:
        X, y = X.to(device), y.to(device)
        optimizer.zero_grad()
        preds = model(X)
        loss = criterion(preds, y)
        loss.backward()
        optimizer.step()
        running_loss += loss.item() * len(X)
    return running_loss / len(loader.dataset)


def _evaluate(
    model: CNNBiLSTM, loader: DataLoader, criterion: nn.Module, device: torch.device
) -> Dict[str, float]:
    model.eval()
    all_preds, all_targets = [], []
    running_loss = 0.0
    with torch.no_grad():
        for X, y in loader:
            X, y = X.to(device), y.to(device)
            preds = model(X)
            loss = criterion(preds, y)
            running_loss += loss.item() * len(X)
            all_preds.append(preds.cpu().numpy())
            all_targets.append(y.cpu().numpy())
    y_true = np.concatenate(all_targets)
    y_pred = np.concatenate(all_preds)
    mse = mean_squared_error(y_true, y_pred)
    return {
        "loss": running_loss / len(loader.dataset),
        "mae": mean_absolute_error(y_true, y_pred),
        "rmse": float(np.sqrt(mse)),
        "r2": r2_score(y_true, y_pred),
    }


def _suggest_model_config(trial: Trial, n_features: int) -> ModelConfig:
    return ModelConfig(
        n_features=n_features,
        cnn_channels=trial.suggest_int("cnn_channels", 16, 128, log=True),
        kernel_size=trial.suggest_int("kernel_size", 2, 5),
        lstm_hidden=trial.suggest_int("lstm_hidden", 32, 256, log=True),
        lstm_layers=trial.suggest_int("lstm_layers", 1, 3),
        dense_units=trial.suggest_int("dense_units", 32, 256, log=True),
        dropout=trial.suggest_float("dropout", 0.05, 0.4),
    )


def run_study(
    frame,
    target_column: str,
    feature_columns=None,
    sequence_config: Optional[SequenceConfig] = None,
    n_trials: int = 20,
    device: Optional[torch.device] = None,
    timeout: Optional[int] = None,
) -> optuna.Study:
    """Execute an Optuna study returning the Study object."""

    cfg = sequence_config or SequenceConfig()
    train_loader, val_loader, _, n_features = prepare_datasets(
        frame, target_column, feature_columns, cfg
    )
    device = device or torch.device("cuda" if torch.cuda.is_available() else "cpu")

    def objective(trial: Trial) -> float:
        model_cfg = _suggest_model_config(trial, n_features)
        model = build_model(model_cfg, device=device)
        lr = trial.suggest_float("lr", 1e-4, 1e-2, log=True)
        weight_decay = trial.suggest_float("weight_decay", 1e-6, 1e-2, log=True)
        optimizer = torch.optim.Adam(model.parameters(), lr=lr, weight_decay=weight_decay)
        criterion = nn.MSELoss()
        patience, best_loss, wait = 5, float("inf"), 0

        for epoch in range(50):
            _train_one_epoch(model, train_loader, criterion, optimizer, device)
            metrics = _evaluate(model, val_loader, criterion, device)
            trial.report(metrics["loss"], step=epoch)
            if trial.should_prune():
                raise optuna.TrialPruned()

            if metrics["loss"] + 1e-6 < best_loss:
                best_loss = metrics["loss"]
                wait = 0
            else:
                wait += 1
            if wait >= patience:
                break
        return best_loss

    study = optuna.create_study(direction="minimize")
    study.optimize(objective, n_trials=n_trials, timeout=timeout)
    return study


def train_with_best_trial(
    frame,
    target_column: str,
    feature_columns=None,
    sequence_config: Optional[SequenceConfig] = None,
    n_trials: int = 20,
    device: Optional[torch.device] = None,
) -> Dict[str, object]:
    """Run Optuna then train/evaluate the best model; returns artifacts."""

    cfg = sequence_config or SequenceConfig()
    study = run_study(
        frame,
        target_column,
        feature_columns=feature_columns,
        sequence_config=cfg,
        n_trials=n_trials,
        device=device,
    )
    best_params = study.best_params
    model_cfg = ModelConfig(
        n_features=len(feature_columns or [c for c in frame.columns if c != target_column]),
        cnn_channels=best_params["cnn_channels"],
        kernel_size=best_params["kernel_size"],
        lstm_hidden=best_params["lstm_hidden"],
        lstm_layers=best_params["lstm_layers"],
        dense_units=best_params["dense_units"],
        dropout=best_params["dropout"],
    )

    train_loader, val_loader, test_loader, _ = prepare_datasets(
        frame, target_column, feature_columns, cfg
    )
    device = device or torch.device("cuda" if torch.cuda.is_available() else "cpu")
    model = build_model(model_cfg, device=device)
    optimizer = torch.optim.Adam(
        model.parameters(), lr=best_params.get("lr", 1e-3), weight_decay=best_params.get("weight_decay", 0.0)
    )
    criterion = nn.MSELoss()

    best_state, best_val = None, float("inf")
    for _ in range(50):
        _train_one_epoch(model, train_loader, criterion, optimizer, device)
        metrics = _evaluate(model, val_loader, criterion, device)
        if metrics["loss"] < best_val:
            best_val = metrics["loss"]
            best_state = model.state_dict()
    if best_state is not None:
        model.load_state_dict(best_state)

    test_metrics = _evaluate(model, test_loader, criterion, device)
    return {
        "study": study,
        "model": model,
        "model_config": model_cfg,
        "metrics": test_metrics,
        "best_params": best_params,
    }


def save_study_results(study: optuna.Study, path: str) -> None:
    """Persist study results to disk."""

    with open(path, "w", encoding="utf-8") as f:
        json.dump({"best_params": study.best_params, "best_value": study.best_value}, f, indent=2)
