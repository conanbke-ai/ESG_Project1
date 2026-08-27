from __future__ import annotations

import gc
from pathlib import Path

from solar_forecast.models.cnn.config import SequenceConfig
from solar_forecast.models.cnn.workflow import train_and_save
from solar_forecast.pipeline.dataset import DatasetLoadPolicy, DatasetRepository
from solar_forecast.pipeline.preprocessing import NumericPreprocessor
from solar_forecast.settings import ModelJobConfig, PROJECT_ROOT


class CnnBiLstmTrainer:
    """Concrete adapter from job configuration to the CNN-BiLSTM workflow."""

    def train(self, config: ModelJobConfig, run_dir: Path, smoke: bool = False) -> dict[str, object]:
        source = Path(str(config.values["input_dataset"]))
        source = source if source.is_absolute() else PROJECT_ROOT / source
        target = str(config.values["target_column"])
        energy_source = config.values.get("energy_source_filter")
        entity_column = config.values.get("entity_column")
        timestamp_column = config.values.get("timestamp_column")
        passthrough = [column for column in (entity_column, timestamp_column) if column]
        feature_columns = list(config.values.get("feature_columns") or [])
        policy = DatasetLoadPolicy(
            chunk_rows=int(config.values.get("csv_chunk_rows", 100_000)),
            memory_limit_mb=int(config.values.get("memory_limit_mb", 1536)),
            numeric_dtype=str(config.values.get("numeric_dtype", "float32")),
        )
        source, raw, load_report = DatasetRepository(source.parent).load_training_frame(
            source,
            columns=[*passthrough, *feature_columns, target],
            numeric_columns=[*feature_columns, target],
            equals_filters={"energy_source": str(energy_source)} if energy_source else None,
            truthy_filter=config.values.get("quality_filter_column"),
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
        sequence = SequenceConfig(
            sequence_length=min(int(config.values.get("sequence_length", 168)), max(2, len(frame) // 5)),
            test_size=float(config.values.get("test_fraction", 0.15)),
            val_size=float(config.values.get("validation_fraction", 0.15)),
            calibration_size=float(config.values.get("calibration_fraction", 0.10)),
            purge_gap_hours=0 if smoke else int(config.values.get("purge_gap_hours", 168)),
            batch_size=int(config.values.get("batch_size", 64)),
            shuffle=bool(config.values.get("shuffle_training_batches", True)),
            append_missing_indicators=bool(
                config.values.get("append_missing_indicators", True)
            ),
        )
        artifacts = train_and_save(
            frame,
            target_column=target,
            feature_columns=prepared.feature_columns,
            sequence_config=sequence,
            n_trials=1 if smoke else int(config.values.get("n_trials", 10)),
            output_dir=str(run_dir),
            use_optuna=False if smoke else bool(config.values.get("use_optuna", True)),
            epochs=1 if smoke else int(config.values.get("epochs", 50)),
            entity_column=entity_column,
            timestamp_column=timestamp_column,
        )
        return {
            "source": str(source), "checkpoint_path": str(artifacts["checkpoint_path"]),
            "features": prepared.feature_columns,
            "metrics": artifacts["metrics"],
            "n_rows": len(frame),
            "temporal_split": artifacts.get("temporal_split"),
            "memory_aware_loading": load_report.to_dict(),
        }


def train(config: ModelJobConfig, *, run_dir: Path, smoke: bool = False) -> dict[str, object]:
    return CnnBiLstmTrainer().train(config, run_dir, smoke)
