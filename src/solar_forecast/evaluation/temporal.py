from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True)
class TemporalSplitConfig:
    """Leakage-safe four-way split shared by every forecasting model."""

    validation_fraction: float = 0.15
    calibration_fraction: float = 0.10
    test_fraction: float = 0.15
    gap_hours: int = 0

    def __post_init__(self) -> None:
        fractions = (
            self.validation_fraction,
            self.calibration_fraction,
            self.test_fraction,
        )
        if any(value <= 0 for value in fractions) or sum(fractions) >= 1:
            raise ValueError(
                "validation_fraction + calibration_fraction + test_fraction "
                "must be positive and less than one"
            )
        if self.gap_hours < 0:
            raise ValueError("gap_hours cannot be negative")

    @property
    def train_fraction(self) -> float:
        return 1.0 - self.validation_fraction - self.calibration_fraction - self.test_fraction


@dataclass(frozen=True)
class TemporalBoundaries:
    train_end: pd.Timestamp
    validation_end: pd.Timestamp
    calibration_end: pd.Timestamp
    gap_hours: int

    def to_dict(self) -> dict[str, object]:
        return {
            "train_end": self.train_end.isoformat(),
            "validation_end": self.validation_end.isoformat(),
            "calibration_end": self.calibration_end.isoformat(),
            "gap_hours": self.gap_hours,
        }


@dataclass(frozen=True)
class TemporalFrameSplits:
    train: pd.DataFrame
    validation: pd.DataFrame
    calibration: pd.DataFrame
    test: pd.DataFrame
    boundaries: TemporalBoundaries


class TemporalSplitter:
    """Apply one set of global timestamp boundaries to all plants.

    The calibration interval is intentionally isolated from model fitting,
    hyperparameter/feature selection, and final Test scoring. It is reserved
    for frozen residual thresholds and other post-model calibration.
    """

    def __init__(self, config: TemporalSplitConfig | None = None):
        self.config = config or TemporalSplitConfig()

    def boundaries(self, timestamps: pd.Series | pd.Index) -> TemporalBoundaries:
        values = pd.Series(timestamps)
        values = pd.to_datetime(values, errors="coerce").dropna().drop_duplicates().sort_values()
        minimum = 8
        if len(values) < minimum:
            raise ValueError(f"At least {minimum} unique timestamps are required for a four-way split")

        n_timestamps = len(values)
        train_end_index = max(1, int(n_timestamps * self.config.train_fraction)) - 1
        validation_end_index = max(
            train_end_index + 1,
            int(n_timestamps * (self.config.train_fraction + self.config.validation_fraction)) - 1,
        )
        calibration_end_index = max(
            validation_end_index + 1,
            int(
                n_timestamps
                * (
                    self.config.train_fraction
                    + self.config.validation_fraction
                    + self.config.calibration_fraction
                )
            )
            - 1,
        )
        if calibration_end_index >= n_timestamps - 1:
            raise ValueError("Temporal fractions leave no timestamps for Test")
        return TemporalBoundaries(
            train_end=pd.Timestamp(values.iloc[train_end_index]),
            validation_end=pd.Timestamp(values.iloc[validation_end_index]),
            calibration_end=pd.Timestamp(values.iloc[calibration_end_index]),
            gap_hours=self.config.gap_hours,
        )

    def labels(
        self,
        timestamps: pd.Series | pd.Index,
        boundaries: TemporalBoundaries,
    ) -> pd.Series:
        values = pd.to_datetime(pd.Series(timestamps), errors="coerce")
        gap = pd.Timedelta(hours=boundaries.gap_hours)
        labels = pd.Series(pd.NA, index=values.index, dtype="string")
        labels.loc[values.le(boundaries.train_end)] = "train"
        labels.loc[
            values.gt(boundaries.train_end + gap) & values.le(boundaries.validation_end)
        ] = "validation"
        labels.loc[
            values.gt(boundaries.validation_end + gap) & values.le(boundaries.calibration_end)
        ] = "calibration"
        labels.loc[values.gt(boundaries.calibration_end + gap)] = "test"
        return labels

    def split_frame(
        self,
        frame: pd.DataFrame,
        timestamp_column: str = "timestamp",
    ) -> TemporalFrameSplits:
        if timestamp_column not in frame:
            raise ValueError(f"Timestamp column is missing: {timestamp_column}")
        prepared = frame.copy()
        prepared[timestamp_column] = pd.to_datetime(prepared[timestamp_column], errors="coerce")
        prepared = prepared.dropna(subset=[timestamp_column]).sort_values(
            timestamp_column, kind="stable"
        )
        boundaries = self.boundaries(prepared[timestamp_column])
        labels = self.labels(prepared[timestamp_column], boundaries)
        partitions = {
            name: prepared.loc[labels.eq(name)].copy()
            for name in ("train", "validation", "calibration", "test")
        }
        empty = [name for name, value in partitions.items() if value.empty]
        if empty:
            raise ValueError(
                f"Temporal split produced empty partitions {empty}; reduce gap_hours or use more data"
            )
        return TemporalFrameSplits(boundaries=boundaries, **partitions)
