"""Memory-bounded dataset utilities for the CNN-BiLSTM pipeline."""
from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
import torch
from torch.utils.data import DataLoader, Dataset

from solar_forecast.evaluation.temporal_split import (
    TemporalBoundaries,
    TemporalSplitConfig,
    TemporalSplitter,
)
from solar_forecast.evaluation.forecast_samples import (
    HISTORICAL_FORECAST_TASK,
    forecast_window_positions,
    validate_observation_frame,
)

from solar_forecast.models.cnn_bilstm.sequence_config import SequenceConfig


class SequenceDataset(Dataset):
    """Compatibility dataset for already-materialized small arrays."""

    def __init__(self, X: np.ndarray, y: np.ndarray):
        self.X = torch.from_numpy(X).float()
        self.y = torch.from_numpy(y).float()

    def __len__(self) -> int:
        return len(self.X)

    def __getitem__(self, idx: int) -> tuple[torch.Tensor, torch.Tensor]:
        return self.X[idx], self.y[idx]


@dataclass
class _EntitySeries:
    features: np.ndarray
    targets: np.ndarray
    target_positions: dict[str, np.ndarray]
    train_feature_rows: np.ndarray
    plant_id: str
    region: str
    plant: str
    timestamps: np.ndarray
    origin_positions: np.ndarray | None = None
    horizon_hours: int | None = None


class LazyWindowSequenceDataset(Dataset):
    """Slice windows on demand instead of materializing repeated 3-D arrays."""

    def __init__(
        self,
        series: Sequence[_EntitySeries],
        split: str,
        sequence_length: int,
    ):
        self.series = [item for item in series if len(item.target_positions[split])]
        self.positions = [item.target_positions[split] for item in self.series]
        counts = np.asarray([len(value) for value in self.positions], dtype=np.int64)
        self.cumulative = np.cumsum(counts)
        self.sequence_length = sequence_length

    def __len__(self) -> int:
        return int(self.cumulative[-1]) if len(self.cumulative) else 0

    def __getitem__(self, idx: int) -> tuple[torch.Tensor, torch.Tensor]:
        if idx < 0:
            idx += len(self)
        series_index = int(np.searchsorted(self.cumulative, idx, side="right"))
        previous = int(self.cumulative[series_index - 1]) if series_index else 0
        target_position = int(self.positions[series_index][idx - previous])
        item = self.series[series_index]
        window_end = (
            int(item.origin_positions[target_position]) + 1
            if item.origin_positions is not None else target_position
        )
        window = item.features[window_end - self.sequence_length : window_end]
        return torch.from_numpy(window).float(), torch.tensor(item.targets[target_position]).float()

    def context_frame(self, start: int, stop: int) -> pd.DataFrame:
        """Materialize metadata only for one sequential prediction batch."""

        records: list[dict[str, object]] = []
        for idx in range(start, min(stop, len(self))):
            series_index = int(np.searchsorted(self.cumulative, idx, side="right"))
            previous = int(self.cumulative[series_index - 1]) if series_index else 0
            target_position = int(self.positions[series_index][idx - previous])
            item = self.series[series_index]
            records.append(
                {
                    "timestamp": item.timestamps[target_position],
                    "plant_id": item.plant_id,
                    "region": item.region,
                    "plant": item.plant,
                }
            )
            if item.origin_positions is not None:
                origin_position = int(item.origin_positions[target_position])
                records[-1].update({
                    "forecast_origin": item.timestamps[origin_position],
                    "horizon_hours": item.horizon_hours,
                    "persistence_pred": float(item.targets[origin_position]),
                })
        return pd.DataFrame.from_records(records)


def _build_sequences(
    values: np.ndarray,
    targets: np.ndarray,
    sequence_length: int,
) -> Tuple[np.ndarray, np.ndarray]:
    """Compatibility helper for small callers; the main path uses lazy windows."""

    sequences: List[np.ndarray] = []
    sequence_targets: List[np.ndarray] = []
    for index in range(sequence_length, len(values)):
        sequences.append(values[index - sequence_length : index])
        sequence_targets.append(targets[index])
    if not sequences:
        raise ValueError("Not enough rows to build a sequence")
    return np.stack(sequences), np.stack(sequence_targets)


@dataclass(frozen=True)
class SequenceLoaders:
    train: DataLoader
    validation: DataLoader
    calibration: DataLoader
    test: DataLoader
    n_features: int
    split_metadata: dict[str, object]


def _position_labels(n_targets: int, cfg: SequenceConfig) -> dict[str, np.ndarray]:
    train_end = max(1, int(n_targets * (1 - cfg.val_size - cfg.calibration_size - cfg.test_size)))
    validation_end = max(
        train_end + 1,
        int(n_targets * (1 - cfg.calibration_size - cfg.test_size)),
    )
    calibration_end = max(validation_end + 1, int(n_targets * (1 - cfg.test_size)))
    if calibration_end >= n_targets:
        raise ValueError("Not enough sequences for train/validation/calibration/test")
    relative = np.arange(n_targets)
    return {
        "train": relative[:train_end],
        "validation": relative[train_end:validation_end],
        "calibration": relative[validation_end:calibration_end],
        "test": relative[calibration_end:],
    }


def _fit_and_transform_training_medians(
    series: Sequence[_EntitySeries],
    feature_columns: Sequence[str],
    *,
    append_missing_indicators: bool,
) -> dict[str, object]:
    # Compute one feature at a time. Concatenating every entity's full Train
    # matrix duplicates O(rows * features) memory before lazy windows even run.
    feature_count = len(feature_columns)
    medians = np.zeros(feature_count, dtype=np.float32)
    all_missing_mask = np.ones(feature_count, dtype=bool)
    for feature_index in range(feature_count):
        feature_parts = [
            item.features[item.train_feature_rows, feature_index]
            for item in series
            if item.train_feature_rows.any()
        ]
        if not feature_parts:
            continue
        values = np.concatenate(feature_parts).astype(np.float32, copy=False)
        values = values[np.isfinite(values)]
        if values.size:
            medians[feature_index] = np.median(values)
            all_missing_mask[feature_index] = False
    all_missing = [
        feature_columns[index]
        for index, is_missing in enumerate(all_missing_mask)
        if is_missing
    ]
    if all_missing and not append_missing_indicators:
        raise ValueError(
            "Training split has no observed values and missing indicators are disabled for: "
            f"{all_missing}"
        )
    medians = np.nan_to_num(medians, nan=0.0).astype(np.float32)
    for item in series:
        # pandas may expose a read-only NumPy view. Imputation owns its buffer
        # so neither read-only views nor writable caller data are modified.
        numeric = item.features.astype(np.float32, copy=True)
        missing = ~np.isfinite(numeric)
        missing_row, missing_feature = np.where(missing)
        numeric[missing_row, missing_feature] = medians[missing_feature]
        if append_missing_indicators:
            transformed = np.empty(
                (len(numeric), numeric.shape[1] * 2), dtype=np.float32
            )
            transformed[:, : numeric.shape[1]] = numeric
            transformed[:, numeric.shape[1] :] = missing
            item.features = transformed
        else:
            item.features = numeric
    return {
        "strategy": "training_split_median_with_missing_indicators",
        "feature_medians": {
            column: float(medians[index]) for index, column in enumerate(feature_columns)
        },
        "all_missing_training_features": all_missing,
        "append_missing_indicators": append_missing_indicators,
        "effective_feature_columns": [
            *feature_columns,
            *(
                [f"{column}__missing" for column in feature_columns]
                if append_missing_indicators
                else []
            ),
        ],
    }


def prepare_dataset_splits(
    frame: pd.DataFrame,
    target_column: str,
    feature_columns: Optional[Sequence[str]] = None,
    config: Optional[SequenceConfig] = None,
    entity_column: Optional[str] = None,
    timestamp_column: Optional[str] = None,
) -> SequenceLoaders:
    """Build one global four-way time split with lazy per-entity windows."""

    cfg = config or SequenceConfig()
    historical = cfg.prediction_task == HISTORICAL_FORECAST_TASK
    if feature_columns is None:
        excluded = {target_column, entity_column, timestamp_column}
        feature_columns = [column for column in frame.columns if column not in excluded]
    feature_columns = list(feature_columns)
    if historical:
        if not entity_column or not timestamp_column:
            raise ValueError("Historical forecasting requires entity_column and timestamp_column")
        if target_column in feature_columns:
            raise ValueError("The future target cannot also be an unshifted feature column")
        frame = validate_observation_frame(
            frame, entity_column=entity_column, timestamp_column=timestamp_column
        )
    if entity_column and not timestamp_column:
        raise ValueError("timestamp_column is required when entity_column is used")
    for column in (entity_column, timestamp_column):
        if column and column not in frame.columns:
            raise ValueError(f"Split column is missing: {column}")

    boundaries: TemporalBoundaries | None = None
    splitter: TemporalSplitter | None = None
    if timestamp_column:
        splitter = TemporalSplitter(
            TemporalSplitConfig(
                validation_fraction=cfg.val_size,
                calibration_fraction=cfg.calibration_size,
                test_fraction=cfg.test_size,
                gap_hours=max(cfg.purge_gap_hours, cfg.forecast_horizon_hours) if historical else cfg.purge_gap_hours,
            )
        )
        boundaries = splitter.boundaries(frame[timestamp_column])
    elif cfg.purge_gap_hours:
        raise ValueError("purge_gap_hours requires a timestamp_column")

    sort_columns = [
        column for column in (entity_column, timestamp_column) if column is not None
    ]
    prepared = frame.sort_values(sort_columns, kind="stable") if sort_columns else frame.copy()
    grouped = (
        prepared.groupby(entity_column, sort=False, observed=True)
        if entity_column
        else [(None, prepared)]
    )
    series: list[_EntitySeries] = []
    for entity_value, group in grouped:
        if len(group) <= cfg.sequence_length:
            continue
        features = group[feature_columns].to_numpy(dtype=np.float32)
        targets = group[target_column].to_numpy(dtype=np.float32)
        origin_positions = None
        if historical:
            absolute_positions, origins = forecast_window_positions(
                group[timestamp_column],
                horizon_hours=cfg.forecast_horizon_hours,
                sequence_length=cfg.sequence_length,
            )
            finite = np.isfinite(targets[absolute_positions]) & np.isfinite(targets[origins])
            absolute_positions, origins = absolute_positions[finite], origins[finite]
            origin_positions = np.full(len(group), -1, dtype=np.int64)
            origin_positions[absolute_positions] = origins
        else:
            absolute_positions = np.arange(cfg.sequence_length, len(group), dtype=np.int64)
        if timestamp_column and splitter and boundaries:
            labels = splitter.labels(
                pd.Series(group[timestamp_column].to_numpy()[absolute_positions]), boundaries
            )
            target_positions = {
                name: absolute_positions[
                    labels.eq(name).fillna(False).to_numpy(dtype=bool)
                ]
                for name in ("train", "validation", "calibration", "test")
            }
            train_rows = pd.to_datetime(group[timestamp_column], errors="coerce").le(
                boundaries.train_end
            ).to_numpy()
            if historical:
                # Fit medians only on observations actually used by training
                # windows, including neither held-out origins nor future labels.
                train_origins = origin_positions[target_positions["train"]]
                usage = np.zeros(len(group) + 1, dtype=np.int64)
                np.add.at(usage, train_origins - cfg.sequence_length + 1, 1)
                np.add.at(usage, train_origins + 1, -1)
                train_rows = np.cumsum(usage[:-1]) > 0
        else:
            relative = _position_labels(len(absolute_positions), cfg)
            target_positions = {
                name: absolute_positions[position] for name, position in relative.items()
            }
            last_train_target = int(target_positions["train"][-1])
            train_rows = np.arange(len(group)) < last_train_target
        series.append(
            _EntitySeries(
                features=features,
                targets=targets,
                target_positions=target_positions,
                train_feature_rows=train_rows,
                plant_id=(
                    str(entity_value)
                    if entity_column and entity_value is not None
                    else "global"
                ),
                region=(
                    str(group["region"].iloc[0])
                    if "region" in group and pd.notna(group["region"].iloc[0])
                    else "unknown"
                ),
                plant=(
                    str(group["plant"].iloc[0])
                    if "plant" in group and pd.notna(group["plant"].iloc[0])
                    else str(entity_value) if entity_value is not None else "global"
                ),
                timestamps=(
                    pd.to_datetime(group[timestamp_column], errors="coerce").to_numpy()
                    if timestamp_column
                    else np.arange(len(group), dtype=np.int64)
                ),
                origin_positions=origin_positions,
                horizon_hours=cfg.forecast_horizon_hours if historical else None,
            )
        )
    if not series:
        raise ValueError("No entity has enough rows to build sequences")
    empty = [
        name
        for name in ("train", "validation", "calibration", "test")
        if not any(len(item.target_positions[name]) for item in series)
    ]
    if empty:
        raise ValueError(
            f"No sequences were assigned to {empty}; reduce purge gap or sequence length"
        )

    preprocessing_state = _fit_and_transform_training_medians(
        series,
        feature_columns,
        append_missing_indicators=cfg.append_missing_indicators,
    )
    datasets = {
        name: LazyWindowSequenceDataset(series, name, cfg.sequence_length)
        for name in ("train", "validation", "calibration", "test")
    }
    loaders = {
        name: DataLoader(
            dataset,
            batch_size=cfg.batch_size,
            shuffle=cfg.shuffle if name == "train" else False,
            num_workers=cfg.num_workers,
        )
        for name, dataset in datasets.items()
    }
    split_metadata: dict[str, object] = {
        "protocol": "global_timestamp_train_validation_calibration_test",
        "evaluation_protocol": (
            "historical_rolling_origin_with_observation_updates"
            if historical else "legacy_previous_row_sequence_estimation"
        ),
        "forecast_horizon_hours": cfg.forecast_horizon_hours if historical else None,
        "sequence_length": cfg.sequence_length,
        "continuous_hourly_windows_required": historical,
        "fractions": {
            "train": 1 - cfg.val_size - cfg.calibration_size - cfg.test_size,
            "validation": cfg.val_size,
            "calibration": cfg.calibration_size,
            "test": cfg.test_size,
        },
        "counts": {name: len(dataset) for name, dataset in datasets.items()},
        "boundaries": boundaries.to_dict() if boundaries else None,
        "test_period": (
            {
                "start": pd.to_datetime(
                    prepared.loc[
                        pd.to_datetime(prepared[timestamp_column], errors="coerce").gt(
                            boundaries.calibration_end
                            + pd.Timedelta(hours=boundaries.gap_hours)
                        ),
                        timestamp_column,
                    ],
                    errors="coerce",
                ).min().isoformat(),
                "end": pd.to_datetime(
                    prepared.loc[
                        pd.to_datetime(prepared[timestamp_column], errors="coerce").gt(
                            boundaries.calibration_end
                            + pd.Timedelta(hours=boundaries.gap_hours)
                        ),
                        timestamp_column,
                    ],
                    errors="coerce",
                ).max().isoformat(),
            }
            if timestamp_column and boundaries
            else None
        ),
        "window_materialization": "lazy_per_batch",
    }
    preprocessing_state["temporal_split"] = split_metadata
    for loader in loaders.values():
        loader.preprocessing_state = preprocessing_state
        loader.split_metadata = split_metadata
    n_features = len(preprocessing_state["effective_feature_columns"])
    return SequenceLoaders(
        train=loaders["train"],
        validation=loaders["validation"],
        calibration=loaders["calibration"],
        test=loaders["test"],
        n_features=n_features,
        split_metadata=split_metadata,
    )


def prepare_datasets(
    frame: pd.DataFrame,
    target_column: str,
    feature_columns: Optional[Sequence[str]] = None,
    config: Optional[SequenceConfig] = None,
    entity_column: Optional[str] = None,
    timestamp_column: Optional[str] = None,
) -> Tuple[DataLoader, DataLoader, DataLoader, int]:
    """Compatibility view returning Train/Validation/Test from four-way splits."""

    splits = prepare_dataset_splits(
        frame,
        target_column,
        feature_columns,
        config,
        entity_column,
        timestamp_column,
    )
    return splits.train, splits.validation, splits.test, splits.n_features


def sequence_from_csv(
    path: str,
    target_column: str,
    feature_columns: Optional[Sequence[str]] = None,
    config: Optional[SequenceConfig] = None,
) -> Tuple[DataLoader, DataLoader, DataLoader, int]:
    frame = pd.read_csv(path)
    return prepare_datasets(frame, target_column, feature_columns, config)


def reconstruct_targets_from_loader(loader: DataLoader) -> Tuple[np.ndarray, np.ndarray]:
    y_true: List[np.ndarray] = []
    for _, target in loader:
        y_true.append(target.numpy())
    values = np.concatenate(y_true)
    return values, values.copy()
