from __future__ import annotations

import gc
import json
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score

from solar_forecast.ensemble.dynamic_gate import normalize_prediction_columns
from solar_forecast.artifacts.manifest import replace_file_atomic
from solar_forecast.evaluation.temporal import TemporalSplitConfig, TemporalSplitter
from solar_forecast.pipeline.dataset import DatasetLoadPolicy, DatasetRepository
from solar_forecast.pipeline.preprocessing import (
    NumericPreprocessor,
    require_model_quality_filter,
)
from solar_forecast.settings import ModelJobConfig, PROJECT_ROOT

from .checkpointing import TrainingCheckpointStore, dataset_signature, stable_signature
from .optimization import OptimizationSettings
from .xgboost_checkpoint import fit_xgboost_resumable
from .xgboost_optimization import XGBoostHyperparameterOptimizer


class XGBoostTrainer:
    """Concrete training strategy that owns XGBoost-specific persistence."""

    def train(self, config: ModelJobConfig, run_dir: Path, smoke: bool = False) -> dict[str, object]:
        target = str(config.values["target_column"])
        energy_source = config.values.get("energy_source_filter")
        feature_columns = list(config.values.get("feature_columns") or [])
        passthrough = ["timestamp", "region", "plant", "plant_id"]
        source, raw, load_report = self._load(
            config,
            columns=[*passthrough, *feature_columns, target],
            numeric_columns=[*feature_columns, target],
            energy_source=str(energy_source) if energy_source else None,
            smoke=smoke,
        )
        prepared = NumericPreprocessor(fill_missing=False).transform(
            raw,
            target,
            feature_columns,
            passthrough_columns=passthrough,
        )
        del raw
        gc.collect()
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
        run_dir.mkdir(parents=True, exist_ok=True)
        checkpoint_store = TrainingCheckpointStore.from_config(config)
        optimization_settings = OptimizationSettings.from_values(
            config.values,
            model="xgboost",
        ).scoped(checkpoint_store.fingerprint)
        optimization_result = None
        if optimization_settings.enabled and not smoke:
            optimization_result = XGBoostHyperparameterOptimizer(
                optimization_settings,
                config.values,
                checkpoint_store=checkpoint_store,
            ).optimize(
                train_frame,
                validation_frame,
                feature_columns=prepared.feature_columns,
                target_column=target,
                artifact_dir=run_dir,
            )
        optimizer_values = config.values.get("optimizer", {})
        params = {
            "n_estimators": (
                20
                if smoke
                else int(
                    optimizer_values.get("trial_max_estimators", 1_500)
                    if optimization_result and isinstance(optimizer_values, dict)
                    else config.values.get("n_estimators", 500)
                )
            ),
            "max_depth": int(config.values.get("max_depth", 8)),
            "learning_rate": float(config.values.get("learning_rate", 0.05)),
            "subsample": float(config.values.get("subsample", 0.9)),
            "colsample_bytree": float(config.values.get("colsample_bytree", 0.9)),
            "tree_method": "hist",
            "max_bin": int(config.values.get("max_bin", 256)),
            "eval_metric": "mae",
            "random_state": int(config.values.get("seed", 42)),
            "n_jobs": int(config.values.get("n_jobs", 4)),
        }
        if optimization_result:
            params.update(optimization_result.best_params)
        if not smoke:
            params["early_stopping_rounds"] = int(
                optimizer_values.get(
                    "early_stopping_rounds",
                    config.values.get("early_stopping_rounds", 10),
                )
                if isinstance(optimizer_values, dict)
                else config.values.get("early_stopping_rounds", 10)
            )
        checkpoint_signature = stable_signature(
            {
                "params": params,
                "feature_columns": prepared.feature_columns,
                "target_column": target,
                "temporal_split": split_metadata,
            }
        )
        checkpoint_stage = f"final_fit_{checkpoint_signature[:20]}"
        fit = fit_xgboost_resumable(
            params,
            train_frame[prepared.feature_columns],
            train_frame[target],
            validation_frame[prepared.feature_columns],
            validation_frame[target],
            store=checkpoint_store,
            stage=checkpoint_stage,
            signature=checkpoint_signature,
            verbose=False,
        )
        model = fit.model
        validation_predicted = model.predict(validation_frame[prepared.feature_columns])
        calibration_predicted = model.predict(calibration_frame[prepared.feature_columns])
        predicted = model.predict(test_frame[prepared.feature_columns])
        model_path = run_dir / "model.json"
        temporary_model = model_path.with_name(f"{model_path.stem}.tmp{model_path.suffix}")
        model.save_model(temporary_model)
        replace_file_atomic(temporary_model, model_path)
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
                    "optimizer": (
                        optimization_result.to_dict()
                        if optimization_result
                        else {
                            "enabled": False,
                            "reason": "smoke_mode" if smoke else "disabled_by_config",
                            "selection_data": "validation_only",
                            "test_usage": "none",
                        }
                    ),
                    "memory_aware_loading": load_report.to_dict(),
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
            split="validation",
        ).to_csv(validation_path, index=False, encoding="utf-8-sig")
        calibration_path = run_dir / "calibration_predictions.csv"
        self._prediction_frame(
            calibration_frame,
            calibration_frame[target].to_numpy(),
            calibration_predicted,
            split="calibration",
        ).to_csv(calibration_path, index=False, encoding="utf-8-sig")
        prediction_path = run_dir / "test_predictions.csv"
        predictions = self._prediction_frame(
            test_frame,
            test_frame[target].to_numpy(),
            predicted,
            split="test",
        )
        predictions.to_csv(prediction_path, index=False, encoding="utf-8-sig")
        metrics = {
            "mae": float(mean_absolute_error(test_frame[target], predicted)),
            "rmse": float(np.sqrt(mean_squared_error(test_frame[target], predicted))),
            "r2": float(r2_score(test_frame[target], predicted)) if len(test_frame) > 1 else float("nan"),
        }
        checkpoint_details = {
            **checkpoint_store.describe(),
            "stage": checkpoint_stage,
            "resumed": fit.resumed,
            "initial_rounds": fit.initial_rounds,
            "completed_rounds": fit.completed_rounds,
            "retained_for_idempotent_resume": True,
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
            "evaluation_contract": {
                "dataset_fingerprint": dataset_signature(source),
                "target": target,
                "target_unit": "MWh",
                "horizon_hours": int(config.values.get("forecast_horizon_hours", 24)),
                "test_start": split_metadata["test_period"]["start"],
                "test_end": split_metadata["test_period"]["end"],
                "prediction_key": ["timestamp", "plant_id"],
                "prediction_schema": "solar-forecast-prediction.v1",
            },
            "optimizer": (
                optimization_result.to_dict()
                if optimization_result
                else {
                    "enabled": False,
                    "reason": "smoke_mode" if smoke else "disabled_by_config",
                }
            ),
            "memory_aware_loading": load_report.to_dict(),
            "checkpoint": checkpoint_details,
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
            "test_period": {
                "start": splits.test["timestamp"].min().isoformat(),
                "end": splits.test["timestamp"].max().isoformat(),
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
    def _load(
        config: ModelJobConfig,
        *,
        columns: list[str],
        numeric_columns: list[str],
        energy_source: str | None,
        smoke: bool,
    ):
        source = Path(str(config.values["input_dataset"]))
        source = source if source.is_absolute() else PROJECT_ROOT / source
        policy = DatasetLoadPolicy(
            chunk_rows=int(config.values.get("csv_chunk_rows", 100_000)),
            memory_limit_mb=int(config.values.get("memory_limit_mb", 1536)),
            numeric_dtype=str(config.values.get("numeric_dtype", "float32")),
        )
        return DatasetRepository(source.parent).load_training_frame(
            source,
            columns=columns,
            numeric_columns=numeric_columns,
            equals_filters={"energy_source": energy_source} if energy_source else None,
            truthy_filter=require_model_quality_filter(config.values),
            row_limit=10_000 if smoke else None,
            policy=policy,
        )

    @staticmethod
    def _prediction_frame(
        context: pd.DataFrame,
        actual: np.ndarray,
        predicted: np.ndarray,
        *,
        split: str,
    ) -> pd.DataFrame:
        normalized = normalize_prediction_columns(context.reset_index(drop=True))
        result = pd.DataFrame({
            "timestamp": normalized.get("timestamp", pd.Series(range(len(actual)))),
            "plant_id": normalized.get("plant_id", "unknown"),
            "region": normalized.get("region", "unknown"),
            "plant": normalized.get("plant", "unknown"),
            "split": split,
            "y_true": actual,
            "y_pred": predicted,
            "xgb_pred": predicted,
        })
        return result


def train(config: ModelJobConfig, *, run_dir: Path, smoke: bool = False) -> dict[str, object]:
    return XGBoostTrainer().train(config, run_dir, smoke)
