from __future__ import annotations

from typing import Iterable, Sequence

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
