"""PyTorch implementation of a CNN-BiLSTM regressor.

The module isolates the model definition so it can be reused across
training, evaluation, Optuna studies, and reinforcement-learning loops.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

import torch
import torch.nn as nn


@dataclass
class ModelConfig:
    """Configuration for the CNN-BiLSTM architecture."""

    n_features: int
    cnn_channels: int = 32
    kernel_size: int = 3
    lstm_hidden: int = 64
    lstm_layers: int = 1
    dense_units: int = 64
    dropout: float = 0.1


class CNNBiLSTM(nn.Module):
    """1D CNN followed by a bidirectional LSTM for sequence regression."""

    def __init__(self, config: ModelConfig):
        super().__init__()
        padding = config.kernel_size // 2
        self.conv = nn.Sequential(
            nn.Conv1d(
                in_channels=config.n_features,
                out_channels=config.cnn_channels,
                kernel_size=config.kernel_size,
                padding=padding,
            ),
            nn.ReLU(),
            nn.BatchNorm1d(config.cnn_channels),
            nn.Dropout(config.dropout),
        )

        self.lstm = nn.LSTM(
            input_size=config.cnn_channels,
            hidden_size=config.lstm_hidden,
            num_layers=config.lstm_layers,
            dropout=config.dropout if config.lstm_layers > 1 else 0.0,
            bidirectional=True,
            batch_first=True,
        )

        self.head = nn.Sequential(
            nn.Linear(config.lstm_hidden * 2, config.dense_units),
            nn.ReLU(),
            nn.Dropout(config.dropout),
            nn.Linear(config.dense_units, 1),
        )

    def forward(self, x: torch.Tensor) -> torch.Tensor:
        # x: (batch, seq_len, features)
        x = x.transpose(1, 2)  # (batch, features, seq_len)
        x = self.conv(x)
        x = x.transpose(1, 2)  # (batch, seq_len, channels)
        output, _ = self.lstm(x)
        last_step = output[:, -1]
        return self.head(last_step).squeeze(-1)


def build_model(config: ModelConfig, device: Optional[torch.device] = None) -> CNNBiLSTM:
    """Helper to build and place the model on a device."""

    model = CNNBiLSTM(config)
    if device is not None:
        model = model.to(device)
    return model
