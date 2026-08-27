from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from xgboost import XGBRegressor

from solar_forecast.ensemble.dynamic_gate import normalize_prediction_columns
from solar_forecast.evaluation.temporal import TemporalSplitConfig, TemporalSplitter
from solar_forecast.pipeline.dataset import DatasetRepository
from solar_forecast.pipeline.preprocessing import NumericPreprocessor
from solar_forecast.settings import ModelJobConfig, PROJECT_ROOT


class XGBoostTrainer:
    """Concrete training strategy that owns XGBoost-specific persistence."""

    def train(self, config: ModelJobConfig, run_dir: Path, smoke: bool = False) -> dict[str, object]:
        source, raw = self._load(config)
        target = str(config.values["target_column"])
        energy_source = config.values.get("energy_source_filter")
        if energy_source:
            if "energy_source" not in raw:
                raise ValueError("Configured energy_source_filter but dataset has no energy_source column")
            raw = raw.loc[raw["energy_source"].eq(str(energy_source))].copy()
        passthrough = [
            column for column in ("timestamp", "region", "plant", "plant_id") if column in raw
        ]
        prepared = NumericPreprocessor(fill_missing=False).transform(
            raw,
            target,
            config.values.get("feature_columns"),
            passthrough_columns=passthrough,
            quality_filter_column=config.values.get("quality_filter_column"),
        )
        frame = prepared.frame
        if "timestamp" in frame:
            frame["timestamp"] = pd.to_datetime(frame["timestamp"], errors="coerce")
            frame = frame.dropna(subset=["timestamp"]).sort_values("timestamp", kind="stable")
        frame = frame.head(512) if smoke else frame
        train_frame, validation_frame, calibration_frame, test_frame, split_metadata = (
            self._chronological_split(
            frame,
            validation_fraction=float(config.values.get("validation_fraction", 0.15)),
            calibration_fraction=float(config.values.get("calibration_fraction", 0.10)),
            test_fraction=float(config.values.get("test_fraction", 0.15)),
            purge_gap_hours=0 if smoke else int(config.values.get("purge_gap_hours", 168)),
            )
        )
        params = {
            "n_estimators": 20 if smoke else int(config.values.get("n_estimators", 500)),
            "max_depth": int(config.values.get("max_depth", 8)),
            "learning_rate": float(config.values.get("learning_rate", 0.05)),
            "subsample": float(config.values.get("subsample", 0.9)),
            "colsample_bytree": float(config.values.get("colsample_bytree", 0.9)),
            "tree_method": "hist",
            "random_state": int(config.values.get("seed", 42)),
            "n_jobs": -1,
        }
        if not smoke:
            params["early_stopping_rounds"] = int(config.values.get("early_stopping_rounds", 10))
        model = XGBRegressor(**params)
        model.fit(
            train_frame[prepared.feature_columns],
            train_frame[target],
            eval_set=[(validation_frame[prepared.feature_columns], validation_frame[target])],
            verbose=False,
        )
        validation_predicted = model.predict(validation_frame[prepared.feature_columns])
        calibration_predicted = model.predict(calibration_frame[prepared.feature_columns])
        predicted = model.predict(test_frame[prepared.feature_columns])
        run_dir.mkdir(parents=True, exist_ok=True)
        model_path = run_dir / "model.json"
        model.save_model(model_path)
        preprocessing_path = run_dir / "preprocessing.json"
        preprocessing_path.write_text(
            json.dumps(
                {
                    "strategy": "xgboost_native_nan_learned_default_direction",
                    "missing_value": "NaN",
                    "feature_missing_fraction": {
                        split: {
                            column: float(partition[column].isna().mean())
                            for column in prepared.feature_columns
                        }
                        for split, partition in {
                            "train": train_frame,
                            "validation": validation_frame,
                            "calibration": calibration_frame,
                            "test": test_frame,
                        }.items()
                    },
                    "quality_filter_column": config.values.get("quality_filter_column"),
                    "temporal_split": split_metadata,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        validation_path = run_dir / "validation_predictions.csv"
        self._prediction_frame(
            validation_frame,
            validation_frame[target].to_numpy(),
            validation_predicted,
        ).to_csv(validation_path, index=False, encoding="utf-8-sig")
        calibration_path = run_dir / "calibration_predictions.csv"
        self._prediction_frame(
            calibration_frame,
            calibration_frame[target].to_numpy(),
            calibration_predicted,
        ).to_csv(calibration_path, index=False, encoding="utf-8-sig")
        prediction_path = run_dir / "test_predictions.csv"
        predictions = self._prediction_frame(test_frame, test_frame[target].to_numpy(), predicted)
        predictions.to_csv(prediction_path, index=False, encoding="utf-8-sig")
        metrics = {
            "mae": float(mean_absolute_error(test_frame[target], predicted)),
            "rmse": float(np.sqrt(mean_squared_error(test_frame[target], predicted))),
            "r2": float(r2_score(test_frame[target], predicted)) if len(test_frame) > 1 else float("nan"),
        }
        return {
            "source": str(source), "model_path": str(model_path),
            "preprocessing_path": str(preprocessing_path),
            "validation_predictions": str(validation_path),
            "calibration_predictions": str(calibration_path),
            "test_predictions": str(prediction_path),
            "features": prepared.feature_columns, "metrics": metrics,
            "n_train": len(train_frame),
            "n_validation": len(validation_frame),
            "n_calibration": len(calibration_frame),
            "n_test": len(test_frame),
            "temporal_split": split_metadata,
        }

    @staticmethod
    def _chronological_split(
        frame: pd.DataFrame,
        *,
        validation_fraction: float,
        calibration_fraction: float,
        test_fraction: float,
        purge_gap_hours: int,
    ) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, dict[str, object]]:
        if "timestamp" not in frame:
            raise ValueError("XGBoost forecasting requires a timestamp column for temporal splitting")
        splitter = TemporalSplitter(
            TemporalSplitConfig(
                validation_fraction=validation_fraction,
                calibration_fraction=calibration_fraction,
                test_fraction=test_fraction,
                gap_hours=purge_gap_hours,
            )
        )
        splits = splitter.split_frame(frame, "timestamp")
        metadata = {
            "protocol": "global_timestamp_train_validation_calibration_test",
            "evaluation_protocol": "rolling_origin_day_ahead_with_observation_updates",
            "boundaries": splits.boundaries.to_dict(),
            "counts": {
                "train": len(splits.train),
                "validation": len(splits.validation),
                "calibration": len(splits.calibration),
                "test": len(splits.test),
            },
        }
        return (
            splits.train,
            splits.validation,
            splits.calibration,
            splits.test,
            metadata,
        )

    @staticmethod
    def _load(config: ModelJobConfig) -> tuple[Path, pd.DataFrame]:
        source = Path(str(config.values["input_dataset"]))
        source = source if source.is_absolute() else PROJECT_ROOT / source
        return DatasetRepository(source.parent).load(source)

    @staticmethod
    def _prediction_frame(context: pd.DataFrame, actual: np.ndarray, predicted: np.ndarray) -> pd.DataFrame:
        normalized = normalize_prediction_columns(context.reset_index(drop=True))
        result = pd.DataFrame({
            "timestamp": normalized.get("timestamp", pd.Series(range(len(actual)))),
            "region": normalized.get("region", "unknown"),
            "plant": normalized.get("plant", "unknown"),
            "y_true": actual,
            "xgb_pred": predicted,
        })
        return result


def train(config: ModelJobConfig, *, run_dir: Path, smoke: bool = False) -> dict[str, object]:
    return XGBoostTrainer().train(config, run_dir, smoke)
