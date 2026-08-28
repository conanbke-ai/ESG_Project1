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

from solar_forecast.models.optimization import OptimizationSettings
from solar_forecast.artifacts.manifest import replace_file_atomic
from solar_forecast.models.checkpointing import (
    CHECKPOINT_CONTRACT,
    TrainingCheckpointStore,
    capture_rng_state,
    dataframe_signature,
    restore_rng_state,
    stable_signature,
)

from .data import SequenceConfig, prepare_dataset_splits, prepare_datasets
from .model import ModelConfig, build_model
from .optimization import _evaluate, _train_one_epoch, save_study_results, train_with_best_trial
from .adaptive import BanditConfig, run_adaptive_training


def _timestamped_dir(base_dir: Path) -> Path:
    base_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    run_dir = base_dir / stamp
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def _write_prediction_artifact(
    model: torch.nn.Module,
    loader,
    path: Path,
    *,
    split: str,
    device: torch.device,
) -> Path:
    """Stream row-aligned CNN predictions without materializing a full table."""

    dataset = loader.dataset
    if not hasattr(dataset, "context_frame"):
        raise TypeError("CNN prediction dataset does not expose row context")
    temporary = path.with_name(path.name + ".part")
    temporary.unlink(missing_ok=True)
    offset = 0
    model.eval()
    with torch.no_grad():
        for features, targets in loader:
            predicted = model(features.to(device)).detach().cpu().numpy().reshape(-1)
            actual = targets.detach().cpu().numpy().reshape(-1)
            context = dataset.context_frame(offset, offset + len(actual))
            if len(context) != len(actual):
                raise ValueError("CNN prediction context is not aligned with model output")
            context["split"] = split
            context["y_true"] = actual
            context["y_pred"] = predicted
            context["cnn_pred"] = predicted
            context.to_csv(
                temporary,
                mode="w" if offset == 0 else "a",
                header=offset == 0,
                index=False,
                encoding="utf-8-sig" if offset == 0 else "utf-8",
            )
            offset += len(actual)
    if offset != len(dataset):
        raise ValueError("CNN prediction artifact row count does not match Test dataset")
    replace_file_atomic(temporary, path)
    return path


def train_and_save(
    frame: pd.DataFrame,
    target_column: str,
    feature_columns: Optional[Sequence[str]] = None,
    sequence_config: Optional[SequenceConfig] = None,
    n_trials: int = 10,
    output_dir: str = "artifacts/models/cnn_bilstm",
    use_optuna: bool = True,
    use_reinforcement: bool = False,
    epochs: int = 50,
    entity_column: Optional[str] = None,
    timestamp_column: Optional[str] = None,
    optimizer_settings: OptimizationSettings | None = None,
    optimizer_trial_epochs: int = 20,
    early_stopping_patience: int = 5,
    optimizer_max_train_sequences: int | None = None,
    optimizer_max_validation_sequences: int | None = None,
    optimizer_timeout_seconds: int | None = None,
    checkpoint_store: TrainingCheckpointStore | None = None,
    checkpoint_root: str | Path | None = None,
    optimizer_storage_path: str | Path | None = None,
) -> Dict[str, object]:
    """Train the model with Optuna and/or reinforcement learning then persist artifacts.

    All artifacts for a run are stored in a timestamped subdirectory within ``output_dir``.
    """

    cfg = sequence_config or SequenceConfig()
    if checkpoint_store is None:
        selected_features = list(feature_columns) if feature_columns is not None else [
            column
            for column in frame.columns
            if column not in {target_column, entity_column, timestamp_column}
        ]
        resolved_checkpoint_root = (
            Path(checkpoint_root)
            if checkpoint_root is not None
            else Path(output_dir) / ".checkpoints"
        )
        checkpoint_store = TrainingCheckpointStore(
            resolved_checkpoint_root,
            model="cnn_bilstm",
            fingerprint=stable_signature(
                {
                    "checkpoint_contract": CHECKPOINT_CONTRACT,
                    "data": dataframe_signature(
                        frame,
                        [
                            *(
                                [entity_column]
                                if entity_column and entity_column in frame
                                else []
                            ),
                            *(
                                [timestamp_column]
                                if timestamp_column and timestamp_column in frame
                                else []
                            ),
                            *selected_features,
                            target_column,
                        ],
                    ),
                    "target_column": target_column,
                    "feature_columns": selected_features,
                    "sequence_config": cfg.__dict__,
                    "epochs": epochs,
                    "use_optuna": use_optuna,
                    "use_reinforcement": use_reinforcement,
                    "optimizer_trial_epochs": optimizer_trial_epochs,
                    "early_stopping_patience": early_stopping_patience,
                    "optimizer_max_train_sequences": optimizer_max_train_sequences,
                    "optimizer_max_validation_sequences": optimizer_max_validation_sequences,
                }
            ),
        )
    if use_optuna and optimizer_settings is None and optimizer_storage_path is not None:
        optimizer_settings = OptimizationSettings(
            enabled=True,
            study_name="cnn_pipeline_v1",
            storage_path=Path(optimizer_storage_path),
            max_trials=n_trials,
            timeout_seconds=optimizer_timeout_seconds,
            seed=42,
            startup_trials=min(5, n_trials),
            pruner_startup_trials=min(5, n_trials),
            pruner_warmup_steps=5,
        ).scoped(checkpoint_store.fingerprint)
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
            epochs=epochs,
            trial_epochs=optimizer_trial_epochs,
            early_stopping_patience=early_stopping_patience,
            maximum_train_sequences=optimizer_max_train_sequences,
            maximum_validation_sequences=optimizer_max_validation_sequences,
            settings=optimizer_settings,
            artifact_dir=run_dir,
            timeout=optimizer_timeout_seconds,
            checkpoint_store=checkpoint_store,
        )
        model, model_cfg = result["model"], result["model_config"]
        study = result["study"]
        checkpoint_stage = result["checkpoint_stage"]
        checkpoint_resumed = bool(result["checkpoint_resumed"])
        evaluation_loaders = result.pop("loaders")
        save_study_results(study, str(run_dir / "optuna_best.json"))
        if optimizer_settings is None:
            save_study_results(study, str(run_dir / "optimization_summary.json"))
            study.trials_dataframe().to_csv(
                run_dir / "optimization_trials.csv",
                index=False,
                encoding="utf-8-sig",
            )
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
        checkpoint_signature = stable_signature(
            {
                "model_config": model_cfg.__dict__,
                "epochs": epochs,
                "train_sequences": len(train_loader.dataset),
                "validation_sequences": len(val_loader.dataset),
                "mode": "fixed_without_optuna",
            }
        )
        checkpoint_stage = f"final_fit_{checkpoint_signature[:20]}"
        start_epoch = 0
        checkpoint_resumed = False
        checkpoint_completed = False
        if checkpoint_store is not None:
            state = checkpoint_store.load_torch(
                checkpoint_stage,
                signature=checkpoint_signature,
                map_location=device,
            )
            if state is not None:
                model.load_state_dict(state["model_state"])
                optimizer.load_state_dict(state["optimizer_state"])
                start_epoch = int(state["next_epoch"])
                restore_rng_state(state.get("rng_state"))
                checkpoint_resumed = True
                checkpoint_completed = bool(state.get("completed", False))
        completed_epoch = start_epoch
        loop_end = start_epoch if checkpoint_completed else epochs
        for epoch in range(start_epoch, loop_end):
            _train_one_epoch(model, train_loader, criterion, optimizer, device)
            completed_epoch = epoch + 1
            if checkpoint_store is not None and (
                completed_epoch % checkpoint_store.cnn_every_epochs == 0
                or completed_epoch == epochs
            ):
                checkpoint_store.save_torch(
                    checkpoint_stage,
                    {
                        "model_state": model.state_dict(),
                        "optimizer_state": optimizer.state_dict(),
                        "next_epoch": completed_epoch,
                        "rng_state": capture_rng_state(),
                    },
                    signature=checkpoint_signature,
                    progress={
                        "next_epoch": completed_epoch,
                        "total_epochs": epochs,
                    },
                    completed=False,
                )
        if checkpoint_store is not None:
            checkpoint_store.save_torch(
                checkpoint_stage,
                {
                    "model_state": model.state_dict(),
                    "optimizer_state": optimizer.state_dict(),
                    "next_epoch": completed_epoch,
                    "rng_state": capture_rng_state(),
                },
                signature=checkpoint_signature,
                progress={"next_epoch": completed_epoch, "total_epochs": epochs},
                completed=True,
            )
        metrics = _evaluate(model, test_loader, criterion, device)
        result = {"model": model, "model_config": model_cfg, "metrics": metrics, "best_params": {}}
        evaluation_loaders = loaders

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
        train_loader, val_loader, adaptive_test_loader, _ = prepare_datasets(
            frame, target_column, feature_columns, cfg, entity_column, timestamp_column
        )
        base_checkpoint_stage = checkpoint_stage
        base_checkpoint_resumed = checkpoint_resumed
        adaptive_config = BanditConfig(actions=[1e-4, 5e-4, 1e-3, 5e-3])
        adaptive_signature = stable_signature(
            {
                "model_config": model_cfg.__dict__,
                "bandit_config": adaptive_config.__dict__,
                "epochs": 30,
                "train_sequences": len(train_loader.dataset),
                "validation_sequences": len(val_loader.dataset),
                "upstream_stage": base_checkpoint_stage,
            }
        )
        checkpoint_stage = f"adaptive_fit_{adaptive_signature[:20]}"
        reinforcement = run_adaptive_training(
            model_cfg,
            train_loader,
            val_loader,
            epochs=30,
            bandit_cfg=adaptive_config,
            checkpoint_store=checkpoint_store,
            checkpoint_stage=checkpoint_stage,
            checkpoint_signature=adaptive_signature,
            initial_model=model,
        )
        model = reinforcement.model
        checkpoint_resumed = reinforcement.checkpoint_resumed
        adaptive_device = next(model.parameters()).device
        result["metrics"] = _evaluate(
            model,
            adaptive_test_loader,
            torch.nn.MSELoss(),
            adaptive_device,
        )
        with open(run_dir / "reinforcement_history.json", "w", encoding="utf-8") as f:
            json.dump(reinforcement.history, f, indent=2)
        with open(run_dir / "bandit_q_values.json", "w", encoding="utf-8") as f:
            json.dump(reinforcement.q_values, f, indent=2)

    checkpoint_path = run_dir / "cnn_bilstm.pt"
    temporary_checkpoint = checkpoint_path.with_name(
        f"{checkpoint_path.stem}.tmp{checkpoint_path.suffix}"
    )
    torch.save(
        {
            "model_state": model.state_dict(),
            "config": model_cfg.__dict__,
            "feature_columns": list(feature_columns or []),
            "preprocessing": preprocessing_state,
        },
        temporary_checkpoint,
    )
    replace_file_atomic(temporary_checkpoint, checkpoint_path)
    with open(run_dir / "metrics.json", "w", encoding="utf-8") as f:
        json.dump(result.get("metrics", {}), f, indent=2)
    with open(run_dir / "best_params.json", "w", encoding="utf-8") as f:
        json.dump(result.get("best_params", {}), f, indent=2)

    prediction_device = next(model.parameters()).device
    prediction_paths = {
        split: _write_prediction_artifact(
            model,
            getattr(evaluation_loaders, split),
            run_dir / f"{split}_predictions.csv",
            split=split,
            device=prediction_device,
        )
        for split in ("validation", "calibration", "test")
    }

    optimizer_artifact = (
        {
            "enabled": True,
            "selection_data": "validation_only",
            "objective_metric": "validation_mae",
            "best_validation_mae": float(study.best_value),
            "best_params": dict(study.best_params),
            "summary_path": str(run_dir / "optimization_summary.json"),
            "trials_path": str(run_dir / "optimization_trials.csv"),
            "test_usage": "none",
        }
        if use_optuna
        else {
            "enabled": False,
            "reason": "disabled_by_config",
        }
    )

    checkpoint_artifact = (
        {
            **checkpoint_store.describe(),
            "stage": checkpoint_stage,
            "upstream_stage": base_checkpoint_stage if use_reinforcement else None,
            "upstream_resumed": (
                base_checkpoint_resumed if use_reinforcement else None
            ),
            "resumed": checkpoint_resumed,
            "retained_for_idempotent_resume": True,
        }
        if checkpoint_store is not None
        else {
            "enabled": False,
            "resume": False,
            "reason": "checkpoint_store_not_configured",
        }
    )
    return {
        "model": model,
        "config": model_cfg,
        "metrics": result.get("metrics", {}),
        "output_dir": str(run_dir),
        "checkpoint_path": str(checkpoint_path),
        "validation_predictions": str(prediction_paths["validation"]),
        "calibration_predictions": str(prediction_paths["calibration"]),
        "test_predictions": str(prediction_paths["test"]),
        "temporal_split": temporal_split,
        "optimizer": optimizer_artifact,
        "checkpoint": checkpoint_artifact,
    }


def load_checkpoint(path: str, device: Optional[torch.device] = None):
    data = torch.load(path, map_location=device or "cpu", weights_only=False)
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
    # Internal resumable states are also .pt files but do not implement the
    # deployable model artifact contract below.
    for checkpoint in Path(checkpoint_dir).rglob("cnn_bilstm.pt"):
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
