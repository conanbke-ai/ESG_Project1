"""예측 정합성 검증과 발전소·지역·전국 회귀 오차 집계."""
from __future__ import annotations

from typing import Any, Iterable, Sequence

import numpy as np
import pandas as pd
from sklearn.metrics import r2_score


REQUIRED_PREDICTION_COLUMNS = {"region", "plant", "y_true", "y_pred"}


def validate_predictions(frame: pd.DataFrame) -> None:
    missing = REQUIRED_PREDICTION_COLUMNS - set(frame.columns)
    if missing:
        raise ValueError(f"Prediction columns are missing: {sorted(missing)}")
    if frame[["y_true", "y_pred"]].isna().any().any():
        raise ValueError("Predictions contain missing y_true or y_pred values")


def calculate_metrics(frame: pd.DataFrame) -> dict[str, float]:
    """Calculate metrics from rows instead of averaging pre-calculated RMSE/R2."""
    validate_predictions(frame)
    error = frame["y_true"].to_numpy(float) - frame["y_pred"].to_numpy(float)
    return {
        "n_samples": int(len(frame)),
        "mae": float(np.mean(np.abs(error))),
        "rmse": float(np.sqrt(np.mean(np.square(error)))),
        "r2": float(r2_score(frame["y_true"], frame["y_pred"])) if len(frame) > 1 else float("nan"),
        "sum_absolute_error": float(np.abs(error).sum()),
        "sum_squared_error": float(np.square(error).sum()),
    }


def _group_metrics(frame: pd.DataFrame, columns: Sequence[str]) -> pd.DataFrame:
    rows = []
    grouper = columns[0] if len(columns) == 1 else list(columns)
    for key, group in frame.groupby(grouper, dropna=False, sort=True):
        keys = (key,) if len(columns) == 1 else key
        rows.append({**dict(zip(columns, keys)), **calculate_metrics(group)})
    return pd.DataFrame(rows)


def aggregate_metrics(frame: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Return exact plant, region and national metrics from aligned predictions."""
    validate_predictions(frame)
    plant = _group_metrics(frame, ["region", "plant"])
    region = _group_metrics(frame, ["region"])
    national = pd.DataFrame([calculate_metrics(frame)])
    national["plant_macro_mae"] = plant["mae"].mean()
    national["plant_macro_rmse"] = plant["rmse"].mean()
    return {"plant": plant, "region": region, "national": national}


def validation_diagnostics(
    y_true,
    y_pred,
    *,
    persistence_pred=None,
    is_daylight=None,
) -> dict[str, Any]:
    """Return JSON-safe Validation diagnostics while keeping MAE as selector.

    The optimizer still minimizes MAE. These diagnostics make large errors,
    systematic bias, daylight-only accuracy, and persistence-baseline skill
    visible without turning model selection into an opaque composite score.
    """

    truth = np.asarray(y_true, dtype=float).reshape(-1)
    prediction = np.asarray(y_pred, dtype=float).reshape(-1)
    if truth.size == 0 or truth.shape != prediction.shape:
        raise ValueError("Validation truth and prediction must be non-empty and aligned")
    if not (np.isfinite(truth).all() and np.isfinite(prediction).all()):
        raise ValueError("Validation truth and prediction must be finite")

    error = prediction - truth
    mae = float(np.mean(np.abs(error)))
    rmse = float(np.sqrt(np.mean(np.square(error))))
    r2 = float(r2_score(truth, prediction)) if truth.size > 1 else float("nan")
    result: dict[str, Any] = {
        "mae_mwh": mae,
        "rmse_mwh": rmse,
        "r2": r2 if np.isfinite(r2) else None,
        "bias_mwh": float(np.mean(error)),
        "rows": int(truth.size),
        "selection_objective": "mae_mwh",
        "daylight_mae_mwh": None,
        "daylight_rows": 0,
        "persistence_mae_mwh": None,
        "persistence_skill_pct": None,
    }

    if is_daylight is not None:
        daylight = np.asarray(is_daylight).reshape(-1)
        if daylight.shape != truth.shape:
            raise ValueError("Daylight mask must align with Validation rows")
        daylight_mask = daylight.astype(bool)
        result["daylight_rows"] = int(daylight_mask.sum())
        if daylight_mask.any():
            result["daylight_mae_mwh"] = float(
                np.mean(np.abs(error[daylight_mask]))
            )

    if persistence_pred is not None:
        persistence = np.asarray(persistence_pred, dtype=float).reshape(-1)
        if persistence.shape != truth.shape:
            raise ValueError("Persistence baseline must align with Validation rows")
        finite = np.isfinite(persistence)
        if finite.any():
            persistence_error = persistence[finite] - truth[finite]
            persistence_mae = float(np.mean(np.abs(persistence_error)))
            model_mae = float(np.mean(np.abs(error[finite])))
            result["persistence_mae_mwh"] = persistence_mae
            if persistence_mae > 0:
                result["persistence_skill_pct"] = float(
                    (1.0 - model_mae / persistence_mae) * 100.0
                )
    return result
