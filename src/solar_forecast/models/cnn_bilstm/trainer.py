"""CNN 모델 JSON 설정을 데이터 로딩과 학습 workflow에 연결."""
from __future__ import annotations

import gc
from pathlib import Path

from solar_forecast.models.cnn_bilstm.sequence_config import SequenceConfig
from solar_forecast.evaluation.temporal_split import calendar_split_for_execution
from solar_forecast.evaluation.forecast_samples import (
    HISTORICAL_FORECAST_TASK,
    forecast_evaluation_contract,
)
from solar_forecast.models.cnn_bilstm.training_workflow import train_cnn_bilstm
from solar_forecast.models.cnn_bilstm.optimization import optimize_cnn_bilstm
from solar_forecast.models.shared.checkpoint_store import TrainingCheckpointStore, dataset_signature
from solar_forecast.models.shared.optuna_study import OptimizationSettings
from solar_forecast.datasets.repository import DatasetLoadPolicy, DatasetRepository
from solar_forecast.datasets.numeric_preprocessor import (
    NumericPreprocessor,
    require_model_quality_filter,
)
from solar_forecast.config_loader import ModelJobConfig, PROJECT_ROOT


class CnnBiLstmTrainer:
    """Concrete adapter from job configuration to the CNN-BiLSTM workflow."""

    def train(self, config: ModelJobConfig, run_dir: Path, smoke: bool = False, selection_only: bool = False) -> dict[str, object]:
        source = Path(str(config.values["input_dataset"]))
        source = source if source.is_absolute() else PROJECT_ROOT / source
        target = str(config.values["target_column"])
        energy_source = config.values.get("energy_source_filter")
        entity_column = config.values.get("entity_column")
        timestamp_column = config.values.get("timestamp_column")
        passthrough = list(
            dict.fromkeys(
                column
                for column in (entity_column, timestamp_column, "region", "plant")
                if column
            )
        )
        feature_columns = list(config.values.get("feature_columns") or [])
        policy = DatasetLoadPolicy(
            chunk_rows=int(config.values.get("csv_chunk_rows", 100_000)),
            memory_limit_mb=int(config.values.get("memory_limit_mb", 1536)),
            numeric_dtype=str(config.values.get("numeric_dtype", "float32")),
        )
        admitted = list(config.values.get("admitted_plant_ids") or [])
        source, raw, load_report = DatasetRepository(source.parent).load_training_frame(
            source,
            columns=[*passthrough, *feature_columns, target],
            numeric_columns=[*feature_columns, target],
            equals_filters={"energy_source": str(energy_source)} if energy_source else None,
            allowed_values_filters={"plant_id": admitted} if admitted else None,
            truthy_filter=require_model_quality_filter(config.values),
            row_limit=10_000 if smoke else None,
            policy=policy,
        )
        prepared = NumericPreprocessor(fill_missing=False).transform(
            raw,
            target,
            feature_columns,
            passthrough_columns=passthrough,
        )
        del raw
        gc.collect()
        if smoke and entity_column:
            first_entity = prepared.frame[entity_column].iloc[0]
            frame = prepared.frame[prepared.frame[entity_column] == first_entity].head(512)
        else:
            frame = prepared.frame.head(512) if smoke else prepared.frame
        task_contract = forecast_evaluation_contract(
            config.values.get("prediction_task"),
            config.values.get("forecast_horizon_hours"),
            legacy_task="previous_row_sequence_estimation",
        )
        historical = task_contract["task"] == HISTORICAL_FORECAST_TASK
        requested_length = int(config.values.get("sequence_length", 168))
        calendar_split, smoke_override = calendar_split_for_execution(config.values, smoke=smoke)
        sequence = SequenceConfig(
            sequence_length=(
                requested_length if historical and not smoke
                else min(requested_length, max(2, len(frame) // 5))
            ),
            test_size=float(config.values.get("test_fraction", 0.15)),
            val_size=float(config.values.get("validation_fraction", 0.15)),
            calibration_size=float(config.values.get("calibration_fraction", 0.10)),
            purge_gap_hours=0 if smoke else int(config.values.get("purge_gap_hours", 168)),
            batch_size=int(config.values.get("batch_size", 64)),
            shuffle=bool(config.values.get("shuffle_training_batches", True)),
            append_missing_indicators=bool(
                config.values.get("append_missing_indicators", True)
            ),
            prediction_task=config.values.get("prediction_task"),
            forecast_horizon_hours=task_contract["horizon_hours"] if historical else 1,
            **calendar_split,
        )
        optimization_settings = OptimizationSettings.from_values(
            config.values,
            model="cnn_bilstm",
        )
        optimizer_values = config.values.get("optimizer", {})
        if not isinstance(optimizer_values, dict):
            raise ValueError("optimizer configuration must be an object")
        use_optuna = optimization_settings.enabled and not smoke
        checkpoint_store = TrainingCheckpointStore.from_config(config)
        optimization_settings = optimization_settings.scoped(
            checkpoint_store.fingerprint
        )
        if selection_only:
            if not use_optuna:
                raise ValueError("selection_only requires enabled CNN-BiLSTM optimization")
            study = optimize_cnn_bilstm(
                frame,
                target_column=target,
                feature_columns=prepared.feature_columns,
                sequence_config=sequence,
                n_trials=optimization_settings.max_trials,
                entity_column=entity_column,
                timestamp_column=timestamp_column,
                trial_epochs=int(optimizer_values.get("trial_epochs", 20)),
                early_stopping_patience=int(
                    optimizer_values.get(
                        "early_stopping_patience",
                        config.values.get("early_stopping_patience", 5),
                    )
                ),
                maximum_train_sequences=optimizer_values.get(
                    "tuning_train_max_sequences", 250_000
                ),
                maximum_validation_sequences=optimizer_values.get(
                    "tuning_validation_max_sequences", 100_000
                ),
                settings=optimization_settings,
                artifact_dir=run_dir,
                checkpoint_store=checkpoint_store,
                optimizer_parameter_space=optimizer_values.get("search_space"),
            )
            return {
                "source": str(source),
                "features": prepared.feature_columns,
                "n_rows": len(frame),
                "evaluation_contract": {
                    "dataset_fingerprint": dataset_signature(source),
                    "target": target,
                    "target_unit": "MWh",
                    **task_contract,
                },
                "optimizer": {
                    "enabled": True,
                    "selection_data": "validation_only",
                    "objective_metric": "validation_mae",
                    "best_validation_mae": float(study.best_value),
                    "best_params": dict(study.best_params),
                    "validation_cohort": dict(
                        study.user_attrs.get(
                            "current_validation_cohort",
                            study.best_trial.user_attrs.get(
                                "validation_cohort",
                                {},
                            ),
                        )
                    ),
                    "summary_path": str(run_dir / "optimization_summary.json"),
                    "trials_path": str(run_dir / "optimization_trials.csv"),
                    "test_usage": "none",
                },
                "memory_aware_loading": load_report.to_dict(),
                "checkpoint": checkpoint_store.describe(),
                "selection_only": True,
            }

        artifacts = train_cnn_bilstm(
            frame,
            target_column=target,
            feature_columns=prepared.feature_columns,
            sequence_config=sequence,
            n_trials=1 if smoke else optimization_settings.max_trials,
            output_dir=str(run_dir),
            use_optuna=use_optuna,
            epochs=1 if smoke else int(config.values.get("epochs", 50)),
            entity_column=entity_column,
            timestamp_column=timestamp_column,
            optimizer_settings=optimization_settings if use_optuna else None,
            optimizer_trial_epochs=int(optimizer_values.get("trial_epochs", 20)),
            early_stopping_patience=int(
                optimizer_values.get(
                    "early_stopping_patience",
                    config.values.get("early_stopping_patience", 5),
                )
            ),
            optimizer_max_train_sequences=optimizer_values.get(
                "tuning_train_max_sequences", 250_000
            ),
            optimizer_max_validation_sequences=optimizer_values.get(
                "tuning_validation_max_sequences", 100_000
            ),
            checkpoint_store=checkpoint_store,
            optimizer_parameter_space=optimizer_values.get("search_space"),
            seed=int(config.values.get("seed", 42)),
        )
        optimizer_artifact = dict(artifacts.get("optimizer", {}))
        if smoke:
            optimizer_artifact = {"enabled": False, "reason": "smoke_mode"}
        temporal_split = artifacts.get("temporal_split") or {}
        if smoke_override is not None:
            temporal_split["smoke_split_override"] = smoke_override
        test_period = temporal_split.get("test_period") or {}
        return {
            "source": str(source), "checkpoint_path": str(artifacts["checkpoint_path"]),
            "validation_predictions": str(artifacts["validation_predictions"]),
            "calibration_predictions": str(artifacts["calibration_predictions"]),
            "test_predictions": str(artifacts["test_predictions"]),
            "features": prepared.feature_columns,
            "metrics": artifacts["metrics"],
            "n_rows": len(frame),
            "admitted_plant_ids": admitted,
            "temporal_split": temporal_split,
            "evaluation_contract": {
                "dataset_fingerprint": dataset_signature(source),
                "target": target,
                "target_unit": "MWh",
                "test_start": test_period.get("start"),
                "test_end": test_period.get("end"),
                **task_contract,
            },
            "optimizer": optimizer_artifact,
            "memory_aware_loading": load_report.to_dict(),
            "checkpoint": artifacts.get("checkpoint", checkpoint_store.describe()),
        }


def train(config: ModelJobConfig, *, run_dir: Path, smoke: bool = False, selection_only: bool = False) -> dict[str, object]:
    return CnnBiLstmTrainer().train(config, run_dir, smoke, selection_only)