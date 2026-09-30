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


def _aligned_optional(values, size: int, name: str, *, dtype=None):
    if values is None:
        return None
    array = np.asarray(values, dtype=dtype).reshape(-1)
    if array.shape != (size,):
        raise ValueError(f"{name} must align with Validation rows")
    return array


def _json_scalar(value):
    if value is None:
        return None
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating,)):
        value = float(value)
    if isinstance(value, float):
        return value if np.isfinite(value) else None
    if isinstance(value, (pd.Timestamp, np.datetime64)):
        return pd.Timestamp(value).isoformat()
    if pd.isna(value):
        return None
    return str(value) if not isinstance(value, (str, int, bool)) else value


def validation_diagnostics(
    y_true,
    y_pred,
    *,
    persistence_pred=None,
    is_daylight=None,
    capacity_mw=None,
    plant_id=None,
    plant=None,
    region=None,
    timestamp=None,
    top_k: int = 20,
) -> dict[str, Any]:
    """Return multi-axis Validation diagnostics while keeping MAE as selector.

    MAE remains the optimizer objective. Additional metrics are diagnostic only:
    tail-error percentiles expose rare large misses; capacity-normalized MAE
    makes differently sized plants comparable; daylight metrics prevent night
    zeroes from masking daytime error; persistence skill verifies that the model
    beats a trivial observed-generation baseline.
    """

    truth = np.asarray(y_true, dtype=float).reshape(-1)
    prediction = np.asarray(y_pred, dtype=float).reshape(-1)
    if truth.size == 0 or truth.shape != prediction.shape:
        raise ValueError("Validation truth and prediction must be non-empty and aligned")
    if not (np.isfinite(truth).all() and np.isfinite(prediction).all()):
        raise ValueError("Validation truth and prediction must be finite")
    if top_k < 0:
        raise ValueError("top_k must be nonnegative")

    size = truth.size
    persistence = _aligned_optional(
        persistence_pred, size, "Persistence baseline", dtype=float
    )
    daylight = _aligned_optional(is_daylight, size, "Daylight mask")
    capacity = _aligned_optional(capacity_mw, size, "Capacity", dtype=float)
    ids = _aligned_optional(plant_id, size, "plant_id")
    names = _aligned_optional(plant, size, "plant")
    regions = _aligned_optional(region, size, "region")
    times = _aligned_optional(timestamp, size, "timestamp")

    error = prediction - truth
    absolute_error = np.abs(error)
    mae = float(np.mean(absolute_error))
    rmse = float(np.sqrt(np.mean(np.square(error))))
    r2 = float(r2_score(truth, prediction)) if size > 1 else float("nan")

    result: dict[str, Any] = {
        "mae_mwh": mae,
        "rmse_mwh": rmse,
        "r2": r2 if np.isfinite(r2) else None,
        "bias_mwh": float(np.mean(error)),
        "abs_error_p95_mwh": float(np.quantile(absolute_error, 0.95)),
        "abs_error_p99_mwh": float(np.quantile(absolute_error, 0.99)),
        "abs_error_max_mwh": float(np.max(absolute_error)),
        "rows": int(size),
        "selection_objective": "mae_mwh",
        "daylight_mae_mwh": None,
        "daylight_rmse_mwh": None,
        "daylight_rows": 0,
        "nmae_capacity_pct": None,
        "capacity_rows": 0,
        "persistence_mae_mwh": None,
        "persistence_skill_pct": None,
        "plant_metrics": [],
        "top_absolute_errors": [],
    }

    daylight_mask = None
    if daylight is not None:
        daylight_mask = daylight.astype(bool)
        result["daylight_rows"] = int(daylight_mask.sum())
        if daylight_mask.any():
            daylight_error = error[daylight_mask]
            result["daylight_mae_mwh"] = float(
                np.mean(np.abs(daylight_error))
            )
            result["daylight_rmse_mwh"] = float(
                np.sqrt(np.mean(np.square(daylight_error)))
            )

    capacity_mask = None
    if capacity is not None:
        capacity_mask = np.isfinite(capacity) & (capacity > 0)
        result["capacity_rows"] = int(capacity_mask.sum())
        if capacity_mask.any():
            result["nmae_capacity_pct"] = float(
                np.mean(
                    absolute_error[capacity_mask] / capacity[capacity_mask]
                )
                * 100.0
            )

    if persistence is not None:
        finite = np.isfinite(persistence)
        if finite.any():
            persistence_error = persistence[finite] - truth[finite]
            persistence_mae = float(np.mean(np.abs(persistence_error)))
            model_mae = float(np.mean(absolute_error[finite]))
            result["persistence_mae_mwh"] = persistence_mae
            if persistence_mae > 0:
                result["persistence_skill_pct"] = float(
                    (1.0 - model_mae / persistence_mae) * 100.0
                )

    if ids is not None:
        frame = pd.DataFrame(
            {
                "plant_id": ids,
                "y_true": truth,
                "y_pred": prediction,
                "error": error,
                "abs_error": absolute_error,
            }
        )
        if names is not None:
            frame["plant"] = names
        if regions is not None:
            frame["region"] = regions
        if capacity is not None:
            frame["capacity_mw"] = capacity

        plant_rows: list[dict[str, Any]] = []
        for key, group in frame.groupby("plant_id", sort=True, dropna=False):
            group_error = group["error"].to_numpy(float)
            group_abs = group["abs_error"].to_numpy(float)
            metrics: dict[str, Any] = {
                "plant_id": _json_scalar(key),
                "plant": (
                    _json_scalar(group["plant"].iloc[0])
                    if "plant" in group else None
                ),
                "region": (
                    _json_scalar(group["region"].iloc[0])
                    if "region" in group else None
                ),
                "rows": int(len(group)),
                "mae_mwh": float(np.mean(group_abs)),
                "rmse_mwh": float(np.sqrt(np.mean(np.square(group_error)))),
                "bias_mwh": float(np.mean(group_error)),
                "abs_error_p95_mwh": float(np.quantile(group_abs, 0.95)),
                "abs_error_max_mwh": float(np.max(group_abs)),
                "nmae_capacity_pct": None,
            }
            if "capacity_mw" in group:
                cap = pd.to_numeric(
                    group["capacity_mw"], errors="coerce"
                ).to_numpy(float)
                valid = np.isfinite(cap) & (cap > 0)
                if valid.any():
                    metrics["nmae_capacity_pct"] = float(
                        np.mean(group_abs[valid] / cap[valid]) * 100.0
                    )
            plant_rows.append(metrics)
        result["plant_metrics"] = sorted(
            plant_rows,
            key=lambda row: (
                -float(row["rmse_mwh"]),
                str(row["plant_id"]),
            ),
        )

    if top_k:
        order = np.argsort(-absolute_error, kind="stable")[: min(top_k, size)]
        top_rows: list[dict[str, Any]] = []
        for index in order:
            row: dict[str, Any] = {
                "rank": len(top_rows) + 1,
                "abs_error_mwh": float(absolute_error[index]),
                "error_mwh": float(error[index]),
                "y_true_mwh": float(truth[index]),
                "y_pred_mwh": float(prediction[index]),
            }
            if ids is not None:
                row["plant_id"] = _json_scalar(ids[index])
            if names is not None:
                row["plant"] = _json_scalar(names[index])
            if regions is not None:
                row["region"] = _json_scalar(regions[index])
            if times is not None:
                row["timestamp"] = _json_scalar(times[index])
            if persistence is not None and np.isfinite(persistence[index]):
                row["persistence_pred_mwh"] = float(persistence[index])
            if capacity is not None and np.isfinite(capacity[index]):
                row["capacity_mw"] = float(capacity[index])
                if capacity[index] > 0:
                    row["abs_error_capacity_pct"] = float(
                        absolute_error[index] / capacity[index] * 100.0
                    )
            if daylight_mask is not None:
                row["is_daylight"] = bool(daylight_mask[index])
            top_rows.append(row)
        result["top_absolute_errors"] = top_rows

    return result

