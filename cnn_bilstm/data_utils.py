"""Dataset utilities for the CNN-BiLSTM pipeline."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
import torch
from sklearn.model_selection import train_test_split
from torch.utils.data import DataLoader, Dataset


@dataclass
class SequenceConfig:
    """Configuration controlling sequence length and splits."""

    sequence_length: int = 24
    test_size: float = 0.2
    val_size: float = 0.2
    batch_size: int = 64
    shuffle: bool = True
    num_workers: int = 0


class SequenceDataset(Dataset):
    def __init__(self, X: np.ndarray, y: np.ndarray):
        self.X = torch.from_numpy(X).float()
        self.y = torch.from_numpy(y).float()

    def __len__(self):
        return len(self.X)

    def __getitem__(self, idx):
        return self.X[idx], self.y[idx]


def _build_sequences(
    values: np.ndarray, targets: np.ndarray, sequence_length: int
) -> Tuple[np.ndarray, np.ndarray]:
    sequences: List[np.ndarray] = []
    sequence_targets: List[np.ndarray] = []
    for i in range(sequence_length, len(values)):
        sequences.append(values[i - sequence_length : i])
        sequence_targets.append(targets[i])
    return np.stack(sequences), np.stack(sequence_targets)


def prepare_datasets(
    frame: pd.DataFrame,
    target_column: str,
    feature_columns: Optional[Sequence[str]] = None,
    config: Optional[SequenceConfig] = None,
) -> Tuple[DataLoader, DataLoader, DataLoader, int]:
    """Create train/val/test dataloaders from a DataFrame."""

    cfg = config or SequenceConfig()
    if feature_columns is None:
        feature_columns = [c for c in frame.columns if c != target_column]

    features = frame[feature_columns].to_numpy(dtype=np.float32)
    targets = frame[target_column].to_numpy(dtype=np.float32)
    X, y = _build_sequences(features, targets, cfg.sequence_length)

    X_train, X_temp, y_train, y_temp = train_test_split(
        X, y, test_size=cfg.test_size + cfg.val_size, shuffle=cfg.shuffle, random_state=42
    )
    relative_val = cfg.val_size / (cfg.test_size + cfg.val_size)
    X_val, X_test, y_val, y_test = train_test_split(
        X_temp, y_temp, test_size=1 - relative_val, shuffle=cfg.shuffle, random_state=42
    )

    train_loader = DataLoader(
        SequenceDataset(X_train, y_train),
        batch_size=cfg.batch_size,
        shuffle=True,
        num_workers=cfg.num_workers,
    )
    val_loader = DataLoader(
        SequenceDataset(X_val, y_val),
        batch_size=cfg.batch_size,
        shuffle=False,
        num_workers=cfg.num_workers,
    )
    test_loader = DataLoader(
        SequenceDataset(X_test, y_test),
        batch_size=cfg.batch_size,
        shuffle=False,
        num_workers=cfg.num_workers,
    )
    return train_loader, val_loader, test_loader, len(feature_columns)


def sequence_from_csv(
    path: str,
    target_column: str,
    feature_columns: Optional[Sequence[str]] = None,
    config: Optional[SequenceConfig] = None,
) -> Tuple[DataLoader, DataLoader, DataLoader, int]:
    """Load a CSV file and construct dataloaders."""

    frame = pd.read_csv(path)
    return prepare_datasets(frame, target_column, feature_columns, config)


def reconstruct_targets_from_loader(loader: DataLoader) -> Tuple[np.ndarray, np.ndarray]:
    """Recover y_true and y_pred arrays from a dataloader with stored predictions."""

    y_true: List[np.ndarray] = []
    for _, y in loader:
        y_true.append(y.numpy())
    return np.concatenate(y_true), np.concatenate(y_true)
