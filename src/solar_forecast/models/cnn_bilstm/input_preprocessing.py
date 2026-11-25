"""Train-only CNN input statistics and frozen, versioned NumPy transforms.

Only one feature's training values are concatenated at a time. Sequence windows
remain lazy, and replay does not depend on Torch or refit any statistics.
"""
from __future__ import annotations

from collections.abc import Mapping, Sequence

import numpy as np

INPUT_PREPROCESSING_CONTRACT = "solar-cnn-input-preprocessing.v2"
STANDARD_STRATEGY = "training_split_median_standard_with_missing_indicators"
LEGACY_STRATEGY = "training_split_median_with_missing_indicators"


def _feature_names(feature_columns: Sequence[str]) -> list[str]:
    names = list(feature_columns)
    if not names or any(not isinstance(name, str) or not name for name in names):
        raise ValueError("CNN feature columns must be nonempty names")
    if len(set(names)) != len(names):
        raise ValueError("CNN feature columns must be unique")
    return names


def _effective_names(names: list[str], indicators: bool) -> list[str]:
    effective = names + ([f"{name}__missing" for name in names] if indicators else [])
    if len(set(effective)) != len(effective):
        raise ValueError("CNN feature names collide with generated missing indicators")
    return effective


def fit_input_preprocessing(
    blocks: Sequence[tuple[np.ndarray, np.ndarray]],
    feature_columns: Sequence[str],
    *,
    append_missing_indicators: bool = True,
) -> dict[str, object]:
    """Fit median then population mean/std on unique input rows used in Train."""
    names = _feature_names(feature_columns)
    effective = _effective_names(names, append_missing_indicators)
    for values, selected in blocks:
        if values.ndim != 2 or values.shape[1] != len(names):
            raise ValueError("CNN training block has a different feature width")
        if selected.shape != (len(values),) or selected.dtype != bool:
            raise ValueError("CNN training row selector must be a boolean row mask")
    if not any(selected.any() for _, selected in blocks):
        raise ValueError("No CNN training feature rows are available")
    medians, means, scales = {}, {}, {}
    all_missing, constant = [], []
    for index, name in enumerate(names):
        # Float64 statistics avoid overflow/cancellation for finite float32 input.
        values = np.concatenate([
            matrix[selected, index] for matrix, selected in blocks if selected.any()
        ]).astype(np.float64)
        finite = np.isfinite(values)
        median = float(np.median(values[finite])) if finite.any() else 0.0
        if not finite.any():
            all_missing.append(name)
        values[~finite] = median
        mean, scale = float(values.mean()), float(values.std(ddof=0))
        if scale == 0.0:
            constant.append(name)
            scale = 1.0
        medians[name], means[name], scales[name] = median, mean, scale
    if all_missing and not append_missing_indicators:
        raise ValueError(
            "Training split has no observed values and missing indicators are disabled for: "
            f"{all_missing}"
        )
    state = {
        "contract": INPUT_PREPROCESSING_CONTRACT,
        "schema_version": 2,
        "strategy": STANDARD_STRATEGY,
        "fit_scope": "unique_feature_rows_used_in_training_windows",
        "feature_columns": names,
        "feature_medians": medians,
        "feature_means": means,
        "feature_scales": scales,
        "standard_deviation_ddof": 0,
        "all_missing_training_features": all_missing,
        "constant_training_features": constant,
        "append_missing_indicators": append_missing_indicators,
        "missing_indicators_scaled": False,
        "effective_feature_columns": effective,
        "target_transform": "identity",
    }
    return validate_input_preprocessing(state, names)


def validate_input_preprocessing(
    state: Mapping[str, object], feature_columns: Sequence[str],
) -> dict[str, object]:
    """Reject incompatible artifacts; recognize old median-only states explicitly."""
    if not isinstance(state, Mapping):
        raise ValueError("A saved CNN preprocessing state is required; refitting is forbidden")
    names = _feature_names(feature_columns)
    result = dict(state)
    strategy = state.get("strategy")
    if strategy == STANDARD_STRATEGY:
        if state.get("contract") != INPUT_PREPROCESSING_CONTRACT or state.get("schema_version") != 2:
            raise ValueError("Unsupported CNN preprocessing contract/schema_version")
        if state.get("feature_columns") != names:
            raise ValueError("Stored CNN feature order differs from requested columns")
        if state.get("missing_indicators_scaled") is not False or state.get("target_transform") != "identity":
            raise ValueError("Unsupported CNN indicator or target transform")
        if state.get("standard_deviation_ddof") != 0:
            raise ValueError("Unsupported CNN standard deviation convention")
        statistics = ("feature_medians", "feature_means", "feature_scales")
    elif strategy == LEGACY_STRATEGY:
        if state.get("schema_version", 1) != 1 or state.get("contract") not in (None, "solar-cnn-input-preprocessing.v1"):
            raise ValueError("Unsupported legacy CNN preprocessing contract/schema_version")
        result.update({"schema_version": 1, "contract": "solar-cnn-input-preprocessing.v1"})
        result["compatibility_mode"] = "legacy_median_only_no_scaling"
        statistics = ("feature_medians",)
    else:
        raise ValueError(f"Unsupported stored CNN preprocessing strategy: {strategy}")
    indicators = state.get("append_missing_indicators")
    if not isinstance(indicators, bool):
        raise ValueError("CNN append_missing_indicators must be boolean")
    if state.get("effective_feature_columns") != _effective_names(names, indicators):
        raise ValueError("Stored CNN effective feature order differs from requested columns")
    for key in statistics:
        mapping = state.get(key)
        if not isinstance(mapping, Mapping) or set(mapping) != set(names):
            raise ValueError(f"Stored CNN {key} must match the complete feature schema")
        try:
            numbers = np.asarray([mapping[name] for name in names], dtype=np.float64)
        except (ValueError, TypeError) as exc:
            raise ValueError(f"Stored CNN {key} must be numeric") from exc
        if not np.isfinite(numbers).all() or (key == "feature_scales" and (numbers <= 0).any()):
            raise ValueError(f"Stored CNN {key} must be finite with positive scales")
    return result


def transform_inputs(
    values: np.ndarray, feature_columns: Sequence[str], state: Mapping[str, object],
) -> np.ndarray:
    """Apply stored statistics without changing caller buffers or fitting anything."""
    names = _feature_names(feature_columns)
    frozen = validate_input_preprocessing(state, names)
    if values.ndim != 2 or values.shape[1] != len(names):
        raise ValueError("CNN input matrix differs from saved feature width")
    # Own a buffer even if pandas provided a read-only view.
    numeric = np.asarray(values, dtype=np.float64).copy()
    missing = ~np.isfinite(numeric)
    medians = np.asarray([frozen["feature_medians"][name] for name in names])
    rows, columns = np.where(missing)
    numeric[rows, columns] = medians[columns]
    if frozen["strategy"] == STANDARD_STRATEGY:
        means = np.asarray([frozen["feature_means"][name] for name in names])
        scales = np.asarray([frozen["feature_scales"][name] for name in names])
        numeric -= means
        numeric /= scales
    if not np.isfinite(numeric).all() or (np.abs(numeric) > np.finfo(np.float32).max).any():
        raise ValueError("CNN transformed input exceeds finite float32 range")
    if frozen["append_missing_indicators"]:
        transformed = np.empty((len(numeric), len(names) * 2), dtype=np.float32)
        transformed[:, :len(names)] = numeric
        transformed[:, len(names):] = missing
        return transformed
    return numeric.astype(np.float32)
