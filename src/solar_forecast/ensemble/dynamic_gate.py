from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

REQUIRED_COLUMNS = {"timestamp", "region", "plant", "y_true", "xgb_pred", "cnn_pred"}
COLUMN_ALIASES = {
    "일시": "timestamp", "datetime": "timestamp", "date_time": "timestamp",
    "지역": "region", "지점": "region", "발전구분": "plant", "발전소": "plant",
    "실제발전량": "y_true", "합산발전량(MWh)": "y_true", "target": "y_true",
    "xgboost_pred": "xgb_pred", "xgb_prediction": "xgb_pred",
    "cnn_bilstm_pred": "cnn_pred", "cnn_prediction": "cnn_pred",
}


@dataclass(frozen=True)
class DynamicGateConfig:
    min_group_samples: int = 8
    min_weight: float = 0.05
    max_weight: float = 0.95

    def __post_init__(self) -> None:
        if self.min_group_samples < 1:
            raise ValueError("min_group_samples must be positive")
        if not 0.0 <= self.min_weight < self.max_weight <= 1.0:
            raise ValueError("weight bounds must satisfy 0 <= min < max <= 1")


class ExplainableDynamicGate:
    """Validation-fitted, context-aware ensemble with auditable decisions."""

    def __init__(self, config: DynamicGateConfig | None = None):
        self.config = config or DynamicGateConfig()
        self.profiles_: pd.DataFrame | None = None

    def fit(self, validation: pd.DataFrame) -> "ExplainableDynamicGate":
        self.profiles_ = _build_profiles(validation, self.config.min_group_samples)
        return self

    def predict(self, test: pd.DataFrame) -> pd.DataFrame:
        if self.profiles_ is None:
            raise RuntimeError("fit() must be called before predict()")
        return _apply_profiles(test, self.profiles_, self.config)

    def fit_predict(self, validation: pd.DataFrame, test: pd.DataFrame) -> pd.DataFrame:
        return self.fit(validation).predict(test)


def normalize_prediction_columns(frame: pd.DataFrame) -> pd.DataFrame:
    """Map known source labels to the hybrid contract and derive hour."""
    result = frame.rename(columns={k: v for k, v in COLUMN_ALIASES.items() if k in frame.columns}).copy()
    if "timestamp" in result.columns:
        result["timestamp"] = pd.to_datetime(result["timestamp"], errors="coerce")
        if "hour" not in result.columns:
            result["hour"] = result["timestamp"].dt.hour
    return result


def validate_aligned_predictions(frame: pd.DataFrame) -> None:
    missing = REQUIRED_COLUMNS - set(frame.columns)
    if missing:
        raise ValueError(
            f"Aligned prediction columns are missing: {sorted(missing)}. "
            "Required: timestamp, region, plant, y_true, xgb_pred, cnn_pred"
        )
    if frame[list(REQUIRED_COLUMNS)].isna().any().any():
        raise ValueError("Aligned predictions contain missing values")
    if frame.duplicated(["timestamp", "region", "plant"]).any():
        raise ValueError("Predictions contain duplicate timestamp/region/plant keys")


def _optimal_convex_weight(y_true: np.ndarray, xgb_pred: np.ndarray, cnn_pred: np.ndarray) -> float:
    delta = xgb_pred - cnn_pred
    denominator = float(np.dot(delta, delta))
    if denominator <= 1e-12:
        return 0.5
    return float(np.clip(np.dot(y_true - cnn_pred, delta) / denominator, 0.0, 1.0))


def fit_region_blend(validation: pd.DataFrame) -> pd.DataFrame:
    """Compatibility baseline: learn one convex weight per region."""
    validation = normalize_prediction_columns(validation)
    validate_aligned_predictions(validation)
    rows = []
    for region, group in validation.groupby("region", sort=True):
        alpha = _optimal_convex_weight(group.y_true.to_numpy(float), group.xgb_pred.to_numpy(float), group.cnn_pred.to_numpy(float))
        rows.append({"region": region, "xgb_weight": alpha, "cnn_weight": 1.0 - alpha, "n_validation": len(group)})
    return pd.DataFrame(rows)


def predict_region_blend(test: pd.DataFrame, weights: pd.DataFrame, fallback_weight: float = 0.5) -> pd.DataFrame:
    """Compatibility baseline: apply frozen regional weights."""
    test = normalize_prediction_columns(test)
    validate_aligned_predictions(test)
    result = test.merge(weights[["region", "xgb_weight"]], on="region", how="left", validate="many_to_one")
    result["xgb_weight"] = result.xgb_weight.fillna(float(fallback_weight)).clip(0.0, 1.0)
    result["cnn_weight"] = 1.0 - result.xgb_weight
    result["hybrid_pred"] = result.xgb_weight * result.xgb_pred + result.cnn_weight * result.cnn_pred
    return result


def _build_profiles(validation: pd.DataFrame, min_group_samples: int) -> pd.DataFrame:
    """Learn a transparent hierarchy of validation-only routing contexts."""
    frame = normalize_prediction_columns(validation)
    validate_aligned_predictions(frame)
    frame["xgb_abs_error"] = (frame.y_true - frame.xgb_pred).abs()
    frame["cnn_abs_error"] = (frame.y_true - frame.cnn_pred).abs()
    frame["model_disagreement"] = (frame.xgb_pred - frame.cnn_pred).abs()
    disagreement_low, disagreement_high = frame["model_disagreement"].quantile([1 / 3, 2 / 3])
    frame["disagreement_band"] = _disagreement_band(
        frame["model_disagreement"], float(disagreement_low), float(disagreement_high)
    )
    frame["time_regime"] = frame["hour"].map(_time_regime)

    def profile(scope: str, group: pd.DataFrame, **context: object) -> dict[str, object]:
        alpha = _optimal_convex_weight(
            group.y_true.to_numpy(float),
            group.xgb_pred.to_numpy(float),
            group.cnn_pred.to_numpy(float),
        )
        return {
            "scope": scope,
            "region": "__all__",
            "plant": "__all__",
            "hour": -1,
            "time_regime": "all",
            "disagreement_band": "all",
            "xgb_mae": group.xgb_abs_error.mean(),
            "cnn_mae": group.cnn_abs_error.mean(),
            "xgb_weight": alpha,
            "n_validation": len(group),
            **context,
        }

    rows = [profile("global", frame)]
    rows.append(
        {
            **profile("metadata", frame),
            "disagreement_low_threshold": float(disagreement_low),
            "disagreement_high_threshold": float(disagreement_high),
        }
    )
    group_specs = (
        ("region", ["region"]),
        ("plant", ["plant"]),
        ("region_hour", ["region", "hour"]),
        ("region_regime_disagreement", ["region", "time_regime", "disagreement_band"]),
        ("plant_hour", ["plant", "hour"]),
    )
    for scope, keys in group_specs:
        for values, group in frame.groupby(keys, sort=True):
            if len(group) < min_group_samples:
                continue
            values = values if isinstance(values, tuple) else (values,)
            context = dict(zip(keys, values))
            if "hour" in context:
                context["hour"] = int(context["hour"])
            rows.append(profile(scope, group, **context))
    return pd.DataFrame(rows)


def _apply_profiles(test: pd.DataFrame, gate: pd.DataFrame, config: DynamicGateConfig) -> pd.DataFrame:
    """Route each row by plant/time/disagreement context and retain its rationale."""
    frame = normalize_prediction_columns(test)
    validate_aligned_predictions(frame)
    global_profile = gate.loc[gate.scope == "global"].iloc[0]
    metadata = gate.loc[gate.scope == "metadata"]
    if metadata.empty:
        disagreement_low = disagreement_high = float("inf")
    else:
        disagreement_low = float(metadata.iloc[0].get("disagreement_low_threshold", float("inf")))
        disagreement_high = float(metadata.iloc[0].get("disagreement_high_threshold", float("inf")))

    result = frame.copy()
    result["model_disagreement"] = (result.xgb_pred - result.cnn_pred).abs()
    result["disagreement_band"] = _disagreement_band(
        result["model_disagreement"], disagreement_low, disagreement_high
    )
    result["time_regime"] = result["hour"].map(_time_regime)
    result["gate_scope"] = "global"
    result["xgb_expected_mae"] = float(global_profile["xgb_mae"])
    result["cnn_expected_mae"] = float(global_profile["cnn_mae"])
    result["xgb_weight"] = float(global_profile.get("xgb_weight", 0.5))
    result["gate_validation_samples"] = int(global_profile["n_validation"])

    # Later scopes are more specific and override only where validation support exists.
    group_specs = (
        ("region", ["region"]),
        ("plant", ["plant"]),
        ("region_hour", ["region", "hour"]),
        ("region_regime_disagreement", ["region", "time_regime", "disagreement_band"]),
        ("plant_hour", ["plant", "hour"]),
    )
    for scope, keys in group_specs:
        profiles = gate.loc[
            gate.scope.eq(scope),
            [*keys, "xgb_mae", "cnn_mae", "xgb_weight", "n_validation"],
        ].rename(
            columns={
                "xgb_mae": "_profile_xgb_mae",
                "cnn_mae": "_profile_cnn_mae",
                "xgb_weight": "_profile_xgb_weight",
                "n_validation": "_profile_n_validation",
            }
        )
        if profiles.empty:
            continue
        result = result.merge(profiles, on=keys, how="left", validate="many_to_one")
        matched = result["_profile_xgb_mae"].notna()
        result.loc[matched, "gate_scope"] = scope
        result.loc[matched, "xgb_expected_mae"] = result.loc[matched, "_profile_xgb_mae"]
        result.loc[matched, "cnn_expected_mae"] = result.loc[matched, "_profile_cnn_mae"]
        result.loc[matched, "xgb_weight"] = result.loc[matched, "_profile_xgb_weight"]
        result.loc[matched, "gate_validation_samples"] = result.loc[
            matched, "_profile_n_validation"
        ].astype(int)
        result = result.drop(
            columns=[
                "_profile_xgb_mae",
                "_profile_cnn_mae",
                "_profile_xgb_weight",
                "_profile_n_validation",
            ]
        )

    result["xgb_weight"] = result["xgb_weight"].clip(config.min_weight, config.max_weight)
    result["cnn_weight"] = 1.0 - result.xgb_weight
    result["selected_model"] = np.where(result.xgb_weight > 0.55, "xgboost", np.where(result.xgb_weight < 0.45, "cnn_bilstm", "balanced"))
    result["decision_reason"] = result.apply(
        lambda r: (
            f"{r.gate_scope} validation evidence (n={int(r.gate_validation_samples)}, "
            f"regime={r.time_regime}, disagreement={r.disagreement_band}): "
            f"XGB MAE={r.xgb_expected_mae:.6g}, CNN MAE={r.cnn_expected_mae:.6g}, "
            f"validation-optimal XGB weight={r.xgb_weight:.3f}"
        ),
        axis=1,
    )
    result["hybrid_pred"] = result.xgb_weight * result.xgb_pred + result.cnn_weight * result.cnn_pred
    return result


def _time_regime(hour: int) -> str:
    if 6 <= int(hour) < 10:
        return "ramp_up"
    if 10 <= int(hour) < 16:
        return "day"
    if 16 <= int(hour) < 20:
        return "ramp_down"
    return "night"


def _disagreement_band(values: pd.Series, low: float, high: float) -> pd.Series:
    return pd.Series(
        np.where(values.le(low), "low", np.where(values.le(high), "medium", "high")),
        index=values.index,
        dtype="string",
    )


def fit_dynamic_gate(validation: pd.DataFrame, min_group_samples: int = 8) -> pd.DataFrame:
    """Functional compatibility wrapper around ExplainableDynamicGate.fit."""
    model = ExplainableDynamicGate(DynamicGateConfig(min_group_samples=min_group_samples)).fit(validation)
    assert model.profiles_ is not None
    return model.profiles_


def predict_dynamic_hybrid(test: pd.DataFrame, gate: pd.DataFrame) -> pd.DataFrame:
    """Functional compatibility wrapper for persisted gate profiles."""
    return _apply_profiles(test, gate, DynamicGateConfig())
