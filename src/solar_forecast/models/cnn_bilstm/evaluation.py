"""CNN-BiLSTM evaluation and its artifact contract."""
from __future__ import annotations
import json
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Sequence
import numpy as np
import pandas as pd
import torch
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from solar_forecast.models.cnn_bilstm.sequence_data import (
    SequenceConfig,
    prepare_dataset_splits,
)
from solar_forecast.models.cnn_bilstm.input_preprocessing import validate_input_preprocessing
from solar_forecast.models.cnn_bilstm.network import (
    CnnBiLstmNetworkConfig,
    build_cnn_bilstm_network,
)
from solar_forecast.models.cnn_bilstm.optimization import evaluate_cnn_bilstm_loader
from solar_forecast.infrastructure.artifact_store import create_run_directory


def _checkpoint_model(data, device: Optional[torch.device] = None):
    cfg = CnnBiLstmNetworkConfig(**data["config"])
    model = build_cnn_bilstm_network(cfg, device=device or torch.device("cpu"))
    model.load_state_dict(data["model_state"])
    model.eval()
    return model, cfg


def load_checkpoint(path: str, device: Optional[torch.device] = None):
    data = torch.load(path, map_location=device or "cpu", weights_only=True)
    return _checkpoint_model(data, device)


def _checkpoint_loaders(
    data, frame, target_column, feature_columns, sequence_config,
    entity_column=None, timestamp_column=None,
):
    """Restore each checkpoint's feature transform and split; never fit on replay."""
    preprocessing = data.get("preprocessing")
    if not isinstance(preprocessing, dict):
        raise ValueError("Checkpoint has no saved preprocessing; replay cannot refit it")
    saved_features = data.get("feature_columns") or list(
        preprocessing.get("effective_feature_columns", [])[:len(preprocessing.get("feature_medians", {}))]
    )
    features = list(feature_columns) if feature_columns is not None else saved_features
    if features != saved_features:
        raise ValueError("Requested feature order differs from checkpoint")
    frozen = validate_input_preprocessing(preprocessing, features)
    saved_config = data.get("sequence_config")
    if saved_config:
        cfg = SequenceConfig(**saved_config)
        if sequence_config is not None:
            for field in saved_config:
                if field not in {"batch_size", "shuffle", "num_workers"} and getattr(sequence_config, field) != getattr(cfg, field):
                    raise ValueError(f"Requested sequence config differs from checkpoint: {field}")
            cfg = sequence_config
    else:
        # Legacy median-only artifacts saved split settings, but not the complete
        # SequenceConfig. Recover available settings rather than re-estimating.
        split = frozen.get("temporal_split") or {}
        fractions = split.get("fractions") or {}
        settings = {
            "sequence_length": split.get("sequence_length", 24),
            "append_missing_indicators": frozen["append_missing_indicators"],
            "val_size": fractions.get("validation", 0.15),
            "calibration_size": fractions.get("calibration", 0.10),
            "test_size": fractions.get("test", 0.15),
        }
        if split.get("evaluation_protocol") == "historical_rolling_origin_with_observation_updates":
            settings.update(prediction_task="historical_forecast", forecast_horizon_hours=split["forecast_horizon_hours"])
        if sequence_config is not None:
            for field, value in settings.items():
                if getattr(sequence_config, field) != value:
                    raise ValueError(f"Requested legacy sequence config differs from checkpoint: {field}")
            cfg = sequence_config
        else:
            cfg = SequenceConfig(**settings)
    for name, requested in (("target_column", target_column), ("entity_column", entity_column), ("timestamp_column", timestamp_column)):
        if name in data and requested is not None and data[name] != requested:
            raise ValueError(f"Requested {name} differs from checkpoint")
    loaders = prepare_dataset_splits(
        frame, target_column, features, cfg,
        entity_column or data.get("entity_column"),
        timestamp_column or data.get("timestamp_column"),
        preprocessing_state=frozen,
    )
    if loaders.n_features != data["config"]["n_features"]:
        raise ValueError("Stored CNN input width differs from saved preprocessing")
    return loaders


def evaluate_model(
    model: torch.nn.Module,
    data_loader,
    device: Optional[torch.device] = None,
) -> Dict[str, float]:
    device = device or torch.device("cuda" if torch.cuda.is_available() else "cpu")
    criterion = torch.nn.MSELoss()
    metrics = evaluate_cnn_bilstm_loader(model.to(device), data_loader, criterion, device)
    return metrics


def compare_checkpoints(
    checkpoint_dir: str,
    frame: pd.DataFrame,
    target_column: str,
    feature_columns: Optional[Sequence[str]] = None,
    sequence_config: Optional[SequenceConfig] = None,
    entity_column: Optional[str] = None,
    timestamp_column: Optional[str] = None,
) -> pd.DataFrame:
    """Load all checkpoints in a directory and compare their metrics."""

    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")

    rows: List[Dict[str, object]] = []
    # Internal resumable states are also .pt files but do not implement the
    # deployable model artifact contract below.
    for checkpoint in Path(checkpoint_dir).rglob("cnn_bilstm.pt"):
        data = torch.load(checkpoint, map_location="cpu", weights_only=True)
        loaders = _checkpoint_loaders(
            data, frame, target_column, feature_columns, sequence_config,
            entity_column, timestamp_column,
        )
        model, model_cfg = _checkpoint_model(data, device=device)
        metrics = evaluate_model(model, loaders.test, device=device)
        rows.append({"checkpoint": checkpoint.name, **metrics, **model_cfg.__dict__})

    return pd.DataFrame(rows).sort_values(by="loss") if rows else pd.DataFrame()


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

    data = torch.load(checkpoint_path, map_location="cpu", weights_only=True)
    loaders = _checkpoint_loaders(
        data, frame, target_column, feature_columns, sequence_config,
        entity_column, timestamp_column,
    )
    calibration_true_all: List[np.ndarray] = []
    calibration_pred_all: List[np.ndarray] = []
    y_true_all: List[np.ndarray] = []
    y_pred_all: List[np.ndarray] = []

    model, _ = _checkpoint_model(data)
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
        run_dir = create_run_directory(Path(output_dir))
        with open(run_dir / "metrics.json", "w", encoding="utf-8") as f:
            json.dump(metrics, f, indent=2)
        anomaly_df.to_csv(run_dir / "anomalies.csv", index=False)
        result["output_dir"] = str(run_dir)

    return result


def save_dataframe(df: pd.DataFrame, path: str) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False)
