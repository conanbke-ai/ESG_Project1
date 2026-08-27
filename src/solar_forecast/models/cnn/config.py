from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class SequenceConfig:
    """Sequence construction and leakage-safe four-way temporal split settings.

    ``shuffle`` applies only to batches inside the already isolated training
    partition. It never randomizes chronological split membership.
    """

    sequence_length: int = 24
    test_size: float = 0.15
    val_size: float = 0.15
    calibration_size: float = 0.10
    purge_gap_hours: int = 0
    batch_size: int = 64
    shuffle: bool = True
    append_missing_indicators: bool = True
    num_workers: int = 0

    def __post_init__(self) -> None:
        if self.sequence_length < 1 or self.batch_size < 1:
            raise ValueError("sequence_length and batch_size must be positive")
        fractions = (self.test_size, self.val_size, self.calibration_size)
        if any(value <= 0 for value in fractions) or sum(fractions) >= 1:
            raise ValueError(
                "test_size + val_size + calibration_size must be positive and less than one"
            )
        if self.purge_gap_hours < 0:
            raise ValueError("purge_gap_hours cannot be negative")
