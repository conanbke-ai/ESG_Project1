"""Optuna study utilities for the CNN-BiLSTM model."""
from __future__ import annotations

import copy
import gc
import json
from pathlib import Path
from typing import Dict, Optional

import numpy as np
import optuna
import torch
import torch.nn as nn
from optuna.trial import Trial
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from torch.utils.data import DataLoader, Subset

from solar_forecast.models.optimization import (
    OptimizationSettings,
    OptunaStudyService,
)

from .data import SequenceConfig, prepare_dataset_splits
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


def _bounded_loader(
    loader: DataLoader,
    maximum_sequences: int | None,
    *,
    shuffle: bool,
) -> DataLoader:
    if maximum_sequences is None or len(loader.dataset) <= maximum_sequences:
        return loader
    if maximum_sequences < 1:
        raise ValueError("optimizer sequence limits must be positive or null")
    indices = np.linspace(
        0,
        len(loader.dataset) - 1,
        num=maximum_sequences,
        dtype=np.int64,
    ).tolist()
    bounded = DataLoader(
        Subset(loader.dataset, indices),
        batch_size=loader.batch_size,
        shuffle=shuffle,
        num_workers=0,
    )
    bounded.preprocessing_state = getattr(loader, "preprocessing_state", None)
    bounded.split_metadata = getattr(loader, "split_metadata", None)
    return bounded


def run_study(
    frame,
    target_column: str,
    feature_columns=None,
    sequence_config: Optional[SequenceConfig] = None,
    n_trials: int = 20,
    device: Optional[torch.device] = None,
    timeout: Optional[int] = None,
    entity_column: Optional[str] = None,
    timestamp_column: Optional[str] = None,
    trial_epochs: int = 20,
    early_stopping_patience: int = 5,
    maximum_train_sequences: int | None = None,
    maximum_validation_sequences: int | None = None,
    settings: OptimizationSettings | None = None,
    artifact_dir: Path | None = None,
) -> optuna.Study:
    """Select architecture and optimizer values from Validation only."""

    cfg = sequence_config or SequenceConfig()
    loaders = prepare_dataset_splits(
        frame, target_column, feature_columns, cfg, entity_column, timestamp_column
    )
    train_loader = _bounded_loader(
        loaders.train,
        maximum_train_sequences,
        shuffle=cfg.shuffle,
    )
    val_loader = _bounded_loader(
        loaders.validation,
        maximum_validation_sequences,
        shuffle=False,
    )
    n_features = loaders.n_features
    device = device or torch.device("cuda" if torch.cuda.is_available() else "cpu")
    if min(trial_epochs, early_stopping_patience) < 1:
        raise ValueError("optimizer trial epochs and patience must be positive")

    def objective(trial: Trial) -> float:
        seed = (settings.seed if settings else 42) + trial.number
        np.random.seed(seed)
        torch.manual_seed(seed)
        model_cfg = _suggest_model_config(trial, n_features)
        model = build_model(model_cfg, device=device)
        lr = trial.suggest_float("lr", 1e-4, 1e-2, log=True)
        weight_decay = trial.suggest_float(
            "weight_decay", 1e-6, 1e-2, log=True
        )
        optimizer = torch.optim.Adam(model.parameters(), lr=lr, weight_decay=weight_decay)
        criterion = nn.MSELoss()
        best_mae, wait = float("inf"), 0

        try:
            for epoch in range(trial_epochs):
                _train_one_epoch(model, train_loader, criterion, optimizer, device)
                metrics = _evaluate(model, val_loader, criterion, device)
                validation_mae = float(metrics["mae"])
                trial.report(validation_mae, step=epoch)
                if trial.should_prune():
                    raise optuna.TrialPruned()

                if validation_mae + 1e-6 < best_mae:
                    best_mae = validation_mae
                    wait = 0
                else:
                    wait += 1
                if wait >= early_stopping_patience:
                    break
            trial.set_user_attr("tuning_train_sequences", len(train_loader.dataset))
            trial.set_user_attr(
                "tuning_validation_sequences", len(val_loader.dataset)
            )
            return best_mae
        finally:
            del model
            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

    if settings:
        if artifact_dir is None:
            raise ValueError("artifact_dir is required for a persistent Optuna study")
        return OptunaStudyService(settings).run(objective, artifact_dir).study
    study = optuna.create_study(
        direction="minimize",
        sampler=optuna.samplers.TPESampler(seed=42),
        pruner=optuna.pruners.MedianPruner(),
    )
    study.optimize(objective, n_trials=n_trials, timeout=timeout, gc_after_trial=True)
    return study


def train_with_best_trial(
    frame,
    target_column: str,
    feature_columns=None,
    sequence_config: Optional[SequenceConfig] = None,
    n_trials: int = 20,
    device: Optional[torch.device] = None,
    entity_column: Optional[str] = None,
    timestamp_column: Optional[str] = None,
    epochs: int = 50,
    trial_epochs: int = 20,
    early_stopping_patience: int = 5,
    maximum_train_sequences: int | None = None,
    maximum_validation_sequences: int | None = None,
    settings: OptimizationSettings | None = None,
    artifact_dir: Path | None = None,
    timeout: Optional[int] = None,
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
        entity_column=entity_column,
        timestamp_column=timestamp_column,
        trial_epochs=trial_epochs,
        early_stopping_patience=early_stopping_patience,
        maximum_train_sequences=maximum_train_sequences,
        maximum_validation_sequences=maximum_validation_sequences,
        settings=settings,
        artifact_dir=artifact_dir,
        timeout=timeout,
    )
    best_params = study.best_params
    loaders = prepare_dataset_splits(
        frame, target_column, feature_columns, cfg, entity_column, timestamp_column
    )
    model_cfg = ModelConfig(
        n_features=loaders.n_features,
        cnn_channels=best_params["cnn_channels"],
        kernel_size=best_params["kernel_size"],
        lstm_hidden=best_params["lstm_hidden"],
        lstm_layers=best_params["lstm_layers"],
        dense_units=best_params["dense_units"],
        dropout=best_params["dropout"],
    )

    train_loader = loaders.train
    val_loader = loaders.validation
    device = device or torch.device("cuda" if torch.cuda.is_available() else "cpu")
    model = build_model(model_cfg, device=device)
    optimizer = torch.optim.Adam(
        model.parameters(), lr=best_params.get("lr", 1e-3), weight_decay=best_params.get("weight_decay", 0.0)
    )
    criterion = nn.MSELoss()

    best_state, best_val, wait = None, float("inf"), 0
    for _ in range(epochs):
        _train_one_epoch(model, train_loader, criterion, optimizer, device)
        metrics = _evaluate(model, val_loader, criterion, device)
        if metrics["mae"] + 1e-6 < best_val:
            best_val = float(metrics["mae"])
            best_state = copy.deepcopy(model.state_dict())
            wait = 0
        else:
            wait += 1
        if wait >= early_stopping_patience:
            break
    if best_state is not None:
        model.load_state_dict(best_state)

    validation_metrics = _evaluate(model, loaders.validation, criterion, device)
    calibration_metrics = _evaluate(model, loaders.calibration, criterion, device)
    test_metrics = _evaluate(model, loaders.test, criterion, device)
    return {
        "study": study,
        "model": model,
        "model_config": model_cfg,
        "metrics": test_metrics,
        "best_params": best_params,
        "validation_metrics": validation_metrics,
        "calibration_metrics": calibration_metrics,
        "preprocessing": getattr(train_loader, "preprocessing_state", None),
    }


def save_study_results(study: optuna.Study, path: str) -> None:
    """Persist study results to disk."""

    with open(path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "selection_data": "validation_only",
                "objective_metric": "validation_mae",
                "best_params": study.best_params,
                "best_value": study.best_value,
                "best_trial_number": study.best_trial.number,
                "test_usage": "none",
            },
            f,
            indent=2,
        )
