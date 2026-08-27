from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from xgboost import XGBRegressor

from solar_forecast.features.engineering import (
    HISTORY_OBSERVATION_FEATURES,
    OBSERVATION_MASK_FEATURES,
    SELECTED_V2_MODEL_FEATURES,
    SOLAR_GEOMETRY_FEATURES,
)
from solar_forecast.pipeline.dataset import DatasetRepository


@dataclass(frozen=True)
class FeatureAblationResult:
    result_path: Path
    manifest_path: Path
    selected_contract: str
    selected_features: tuple[str, ...]
    fold_result_path: Path


@dataclass(frozen=True)
class RollingOriginFold:
    number: int
    train_end: pd.Timestamp
    validation_start: pd.Timestamp
    validation_end: pd.Timestamp

    def to_dict(self) -> dict[str, object]:
        return {
            "fold": self.number,
            "train_end": self.train_end.isoformat(),
            "validation_start": self.validation_start.isoformat(),
            "validation_end": self.validation_end.isoformat(),
        }


class FeatureAblationService:
    """Select features with purged rolling-origin folds before Calibration/Test."""

    def __init__(self, *, seed: int = 42, n_estimators: int = 300):
        self.seed = seed
        self.n_estimators = n_estimators

    def run(
        self,
        dataset_path: Path,
        output_dir: Path,
        *,
        n_splits: int = 3,
        validation_window_hours: int = 2_160,
        calibration_fraction: float = 0.10,
        test_fraction: float = 0.15,
        gap_hours: int = 168,
    ) -> FeatureAblationResult:
        _, frame = DatasetRepository(Path(dataset_path).parent).load(Path(dataset_path))
        frame["timestamp"] = pd.to_datetime(frame["timestamp"], errors="coerce")
        frame = frame.dropna(subset=["timestamp", "generation_mwh"])
        if "energy_source" in frame:
            frame = frame.loc[frame["energy_source"].eq("solar")]
        if "quality_train_eligible" in frame:
            frame = frame.loc[self._as_bool(frame["quality_train_eligible"])]
        folds, reserved = self._rolling_folds(
            frame["timestamp"],
            n_splits=n_splits,
            validation_window_hours=validation_window_hours,
            calibration_fraction=calibration_fraction,
            test_fraction=test_fraction,
            gap_hours=gap_hours,
        )

        contracts = {
            "selected_v2_day_ahead": list(SELECTED_V2_MODEL_FEATURES),
            "selected_v3_physics_aware_day_ahead": [
                *SELECTED_V2_MODEL_FEATURES,
                *SOLAR_GEOMETRY_FEATURES,
            ],
            "selected_v2_plus_observation_masks": [
                *SELECTED_V2_MODEL_FEATURES,
                *OBSERVATION_MASK_FEATURES,
            ],
            "selected_v3_quality_physics_aware_day_ahead": [
                *SELECTED_V2_MODEL_FEATURES,
                *SOLAR_GEOMETRY_FEATURES,
                *OBSERVATION_MASK_FEATURES,
            ],
            "selected_v4_missingness_aware_day_ahead": [
                *SELECTED_V2_MODEL_FEATURES,
                *SOLAR_GEOMETRY_FEATURES,
                *OBSERVATION_MASK_FEATURES,
                *HISTORY_OBSERVATION_FEATURES,
            ],
        }
        rows: list[dict[str, object]] = []
        for name, features in contracts.items():
            missing = sorted(set(features) - set(frame.columns))
            if missing:
                raise ValueError(f"Contract {name} columns are missing: {missing}")
            for fold in folds:
                train = frame.loc[frame["timestamp"].le(fold.train_end)].copy()
                validation = frame.loc[
                    frame["timestamp"].between(
                        fold.validation_start,
                        fold.validation_end,
                        inclusive="both",
                    )
                ].copy()
                if train.empty or validation.empty:
                    raise ValueError(f"Fold {fold.number} has an empty partition")
                model = XGBRegressor(
                    n_estimators=self.n_estimators,
                    max_depth=8,
                    learning_rate=0.05,
                    subsample=0.9,
                    colsample_bytree=0.9,
                    tree_method="hist",
                    random_state=self.seed,
                    n_jobs=-1,
                    missing=np.nan,
                )
                model.fit(train[features], train["generation_mwh"])
                predicted = model.predict(validation[features])
                actual = validation["generation_mwh"].to_numpy(float)
                daylight = validation.get(
                    "is_daylight", pd.Series(1, index=validation.index)
                ).eq(1)
                rows.append(
                    {
                        "contract": name,
                        "features": len(features),
                        "fold": fold.number,
                        "train_end": fold.train_end,
                        "validation_start": fold.validation_start,
                        "validation_end": fold.validation_end,
                        "mae": float(mean_absolute_error(actual, predicted)),
                        "rmse": float(np.sqrt(mean_squared_error(actual, predicted))),
                        "r2": float(r2_score(actual, predicted)),
                        "daylight_mae": float(
                            mean_absolute_error(actual[daylight], predicted[daylight])
                        ),
                        "n_train": len(train),
                        "n_validation": len(validation),
                    }
                )

        fold_results = pd.DataFrame(rows)
        results = (
            fold_results.groupby(["contract", "features"], as_index=False)
            .agg(
                mean_mae=("mae", "mean"),
                std_mae=("mae", "std"),
                worst_fold_mae=("mae", "max"),
                mean_rmse=("rmse", "mean"),
                mean_r2=("r2", "mean"),
                mean_daylight_mae=("daylight_mae", "mean"),
                folds=("fold", "nunique"),
            )
            .sort_values(
                ["mean_mae", "worst_fold_mae", "mean_daylight_mae"], kind="stable"
            )
        )
        selected_name = str(results.iloc[0]["contract"])
        selected_features = tuple(contracts[selected_name])
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        result_path = output_dir / "feature_ablation.csv"
        fold_result_path = output_dir / "feature_ablation_folds.csv"
        results.to_csv(result_path, index=False, encoding="utf-8-sig")
        fold_results.to_csv(fold_result_path, index=False, encoding="utf-8-sig")
        manifest = {
            "created_at": datetime.now().isoformat(),
            "dataset": str(dataset_path),
            "protocol": "purged_expanding_rolling_origin",
            "selection_metric": "mean_mae_then_worst_fold_mae",
            "folds": [fold.to_dict() for fold in folds],
            "reserved_intervals": reserved,
            "gap_hours": gap_hours,
            "missing_value_policy": "xgboost_native_nan",
            "test_usage": "none; Calibration and Test intervals are reserved after feature selection",
            "selected_contract": selected_name,
            "selected_features": list(selected_features),
            "summary_result": str(result_path),
            "fold_result": str(fold_result_path),
            "seed": self.seed,
            "n_estimators": self.n_estimators,
        }
        manifest_path = output_dir / "feature_ablation_manifest.json"
        manifest_path.write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return FeatureAblationResult(
            result_path,
            manifest_path,
            selected_name,
            selected_features,
            fold_result_path,
        )

    @staticmethod
    def _rolling_folds(
        timestamps: pd.Series,
        *,
        n_splits: int,
        validation_window_hours: int,
        calibration_fraction: float,
        test_fraction: float,
        gap_hours: int,
    ) -> tuple[list[RollingOriginFold], dict[str, object]]:
        if n_splits < 2 or validation_window_hours < 1:
            raise ValueError("n_splits must be >=2 and validation_window_hours must be positive")
        if (
            calibration_fraction <= 0
            or test_fraction <= 0
            or calibration_fraction + test_fraction >= 0.5
        ):
            raise ValueError(
                "Reserved calibration/test fractions must be positive and sum below 0.5"
            )
        unique = pd.Series(pd.to_datetime(timestamps, errors="coerce")).dropna()
        unique = unique.drop_duplicates().sort_values().reset_index(drop=True)
        calibration_count = max(1, int(len(unique) * calibration_fraction))
        test_count = max(1, int(len(unique) * test_fraction))
        selection_end = len(unique) - calibration_count - test_count
        required = n_splits * validation_window_hours + gap_hours + 1
        if selection_end < required:
            raise ValueError(
                "Not enough pre-Calibration/Test history for requested rolling-origin folds"
            )
        folds: list[RollingOriginFold] = []
        first_validation_start = selection_end - n_splits * validation_window_hours
        for number in range(n_splits):
            validation_start_index = first_validation_start + number * validation_window_hours
            validation_end_index = validation_start_index + validation_window_hours - 1
            train_end_index = validation_start_index - gap_hours - 1
            folds.append(
                RollingOriginFold(
                    number=number + 1,
                    train_end=pd.Timestamp(unique.iloc[train_end_index]),
                    validation_start=pd.Timestamp(unique.iloc[validation_start_index]),
                    validation_end=pd.Timestamp(unique.iloc[validation_end_index]),
                )
            )
        calibration_start = pd.Timestamp(unique.iloc[selection_end])
        test_start = pd.Timestamp(unique.iloc[len(unique) - test_count])
        reserved = {
            "calibration_start": calibration_start.isoformat(),
            "test_start": test_start.isoformat(),
            "dataset_end": pd.Timestamp(unique.iloc[-1]).isoformat(),
            "calibration_fraction": calibration_fraction,
            "test_fraction": test_fraction,
        }
        return folds, reserved

    @staticmethod
    def _as_bool(series: pd.Series) -> pd.Series:
        if pd.api.types.is_bool_dtype(series):
            return series.fillna(False)
        return series.astype(str).str.strip().str.lower().map(
            {"true": True, "1": True, "yes": True, "false": False, "0": False, "no": False}
        ).fillna(False)
