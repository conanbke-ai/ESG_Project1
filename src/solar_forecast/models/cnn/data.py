"""Memory-bounded dataset utilities for the CNN-BiLSTM pipeline."""
from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
import torch
from torch.utils.data import DataLoader, Dataset

from solar_forecast.evaluation.temporal import (
    TemporalBoundaries,
    TemporalSplitConfig,
    TemporalSplitter,
)

from .config import SequenceConfig


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
        window = item.features[target_position - self.sequence_length : target_position]
        return torch.from_numpy(window).float(), torch.tensor(item.targets[target_position]).float()


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
    training_rows = np.concatenate(
        [item.features[item.train_feature_rows] for item in series], axis=0
    )
    training_rows = np.where(np.isfinite(training_rows), training_rows, np.nan)
    all_missing_mask = np.isnan(training_rows).all(axis=0)
    medians = np.zeros(training_rows.shape[1], dtype=np.float32)
    medians[~all_missing_mask] = np.nanmedian(
        training_rows[:, ~all_missing_mask], axis=0
    )
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
        numeric = np.where(np.isfinite(item.features), item.features, np.nan).astype(np.float32)
        indicator = np.isnan(numeric).astype(np.float32)
        missing_row, missing_feature = np.where(np.isnan(numeric))
        numeric[missing_row, missing_feature] = medians[missing_feature]
        item.features = (
            np.concatenate([numeric, indicator], axis=-1)
            if append_missing_indicators
            else numeric
        )
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
    if feature_columns is None:
        excluded = {target_column, entity_column, timestamp_column}
        feature_columns = [column for column in frame.columns if column not in excluded]
    feature_columns = list(feature_columns)
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
                gap_hours=cfg.purge_gap_hours,
            )
        )
        boundaries = splitter.boundaries(frame[timestamp_column])
    elif cfg.purge_gap_hours:
        raise ValueError("purge_gap_hours requires a timestamp_column")

    sort_columns = [
        column for column in (entity_column, timestamp_column) if column is not None
    ]
    prepared = frame.sort_values(sort_columns, kind="stable") if sort_columns else frame.copy()
    grouped = prepared.groupby(entity_column, sort=False) if entity_column else [(None, prepared)]
    series: list[_EntitySeries] = []
    for _, group in grouped:
        if len(group) <= cfg.sequence_length:
            continue
        features = group[feature_columns].to_numpy(dtype=np.float32)
        targets = group[target_column].to_numpy(dtype=np.float32)
        absolute_positions = np.arange(cfg.sequence_length, len(group), dtype=np.int64)
        if timestamp_column and splitter and boundaries:
            labels = splitter.labels(
                pd.Series(group[timestamp_column].to_numpy()[cfg.sequence_length:]), boundaries
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
        "evaluation_protocol": "rolling_origin_day_ahead_with_observation_updates",
        "fractions": {
            "train": 1 - cfg.val_size - cfg.calibration_size - cfg.test_size,
            "validation": cfg.val_size,
            "calibration": cfg.calibration_size,
            "test": cfg.test_size,
        },
        "counts": {name: len(dataset) for name, dataset in datasets.items()},
        "boundaries": boundaries.to_dict() if boundaries else None,
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
