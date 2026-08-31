from __future__ import annotations

from dataclasses import dataclass
from typing import Mapping, Optional, Sequence

import numpy as np
import pandas as pd


REQUIRED_MODEL_QUALITY_FILTER = "quality_train_eligible"


def require_model_quality_filter(values: Mapping[str, object]) -> str:
    """Fail closed when a forecasting job tries to bypass the Gold quality gate."""

    configured = str(values.get("quality_filter_column") or "").strip()
    if configured != REQUIRED_MODEL_QUALITY_FILTER:
        raise ValueError(
            "Forecast model jobs must set quality_filter_column="
            f"{REQUIRED_MODEL_QUALITY_FILTER!r}; the filter cannot be disabled or replaced"
        )
    return configured


@dataclass(frozen=True)
class PreprocessResult:
    frame: pd.DataFrame
    feature_columns: list[str]
    dropped_rows: int


class NumericPreprocessor:
    """Stateless preprocessing policy with an explicit feature contract."""

    def __init__(self, *, fill_missing: bool = True):
        self.fill_missing = fill_missing

    def transform(
        self,
        frame: pd.DataFrame,
        target_column: str,
        feature_columns: Optional[Sequence[str]] = None,
        passthrough_columns: Optional[Sequence[str]] = None,
        quality_filter_column: str | None = None,
    ) -> PreprocessResult:
        return preprocess_dataset(
            frame,
            target_column,
            feature_columns,
            passthrough_columns,
            fill_missing=self.fill_missing,
            quality_filter_column=quality_filter_column,
        )


class TrainFittedMedianImputer:
    """Fit feature medians on training rows only and reuse them unchanged."""

    def __init__(self, medians: Mapping[str, float] | None = None):
        self.medians_: dict[str, float] | None = dict(medians) if medians is not None else None

    def fit(self, frame: pd.DataFrame, feature_columns: Sequence[str]) -> "TrainFittedMedianImputer":
        numeric = frame[list(feature_columns)].apply(pd.to_numeric, errors="coerce")
        numeric = numeric.replace([np.inf, -np.inf], np.nan)
        medians = numeric.median(axis=0, skipna=True)
        missing = medians[medians.isna()].index.tolist()
        if missing:
            raise ValueError(f"Training split has no observed values for features: {missing}")
        self.medians_ = {column: float(medians[column]) for column in feature_columns}
        return self

    def transform(self, frame: pd.DataFrame, feature_columns: Sequence[str]) -> pd.DataFrame:
        if self.medians_ is None:
            raise RuntimeError("fit() must be called before transform()")
        result = frame.copy()
        for column in feature_columns:
            result[column] = pd.to_numeric(result[column], errors="coerce").replace(
                [np.inf, -np.inf], np.nan
            )
            result[column] = result[column].fillna(self.medians_[column])
        return result

    def to_dict(self) -> dict[str, float]:
        if self.medians_ is None:
            raise RuntimeError("fit() must be called before serialization")
        return dict(self.medians_)


def preprocess_dataset(
    frame: pd.DataFrame,
    target_column: str,
    feature_columns: Optional[Sequence[str]] = None,
    passthrough_columns: Optional[Sequence[str]] = None,
    *,
    fill_missing: bool = True,
    quality_filter_column: str | None = None,
) -> PreprocessResult:
    """Select numeric fields and optionally apply legacy all-data median filling."""
    if target_column not in frame.columns:
        raise ValueError(f"Target column '{target_column}' is missing. Available: {list(frame.columns)}")

    features = list(feature_columns) if feature_columns else [
        column for column in frame.columns if column != target_column and pd.api.types.is_numeric_dtype(frame[column])
    ]
    missing = [column for column in features if column not in frame.columns]
    if missing:
        raise ValueError(f"Feature columns are missing: {missing}")
    if not features:
        raise ValueError("No numeric feature columns were found. Pass --features explicitly.")

    passthrough = list(passthrough_columns or [])
    missing_passthrough = [column for column in passthrough if column not in frame.columns]
    if missing_passthrough:
        raise ValueError(f"Passthrough columns are missing: {missing_passthrough}")

    source = frame
    if quality_filter_column:
        if quality_filter_column not in frame:
            raise ValueError(f"Quality filter column is missing: {quality_filter_column}")
        eligible = frame[quality_filter_column]
        if not pd.api.types.is_bool_dtype(eligible):
            eligible = eligible.astype(str).str.strip().str.lower().map(
                {"true": True, "1": True, "yes": True, "false": False, "0": False, "no": False}
            )
        source = frame.loc[eligible.fillna(False)].copy()

    selected = source[list(dict.fromkeys([*passthrough, *features, target_column]))].copy()
    for column in [*features, target_column]:
        selected[column] = pd.to_numeric(selected[column], errors="coerce")
    before = len(selected)
    selected = selected.dropna(subset=[target_column])
    selected[features] = selected[features].replace([np.inf, -np.inf], np.nan)
    if fill_missing:
        selected[features] = selected[features].fillna(selected[features].median(numeric_only=True))
        selected = selected.dropna(subset=features)
    return PreprocessResult(selected.reset_index(drop=True), features, len(frame) - len(selected))
