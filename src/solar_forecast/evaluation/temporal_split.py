"""Train·Validation·Calibration·Test의 시간 경계와 purge gap 생성."""
from __future__ import annotations

from dataclasses import dataclass
import math
import re
from typing import Mapping
import warnings

import pandas as pd


CALENDAR_FIELDS = ("train_end", "validation_end", "calibration_end", "test_end")


def _naive_timestamp(value: object, field: str) -> pd.Timestamp:
    """Require an explicit local timestamp; never silently convert time zones."""
    if not isinstance(value, (str, pd.Timestamp)):
        raise ValueError(f"{field} must be an ISO local timestamp string")
    if isinstance(value, str) and not re.fullmatch(
        r"\d{4}-\d{2}-\d{2}(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d{1,9})?)?)?", value
    ):
        raise ValueError(f"{field} must be a fixed ISO timezone-naive local timestamp; relative dates and offsets are not allowed")
    try:
        parsed = pd.Timestamp(value)
    except (ValueError, TypeError) as exc:
        raise ValueError(f"{field} must be a valid local timestamp") from exc
    if pd.isna(parsed) or parsed.tzinfo is not None:
        raise ValueError(f"{field} must be timezone-naive (dataset local time, KST)")
    return parsed


@dataclass(frozen=True)
class TemporalSplitConfig:
    """Global four-way split, using fractions or four inclusive local timestamps.

    Calendar end dates are inclusive timestamps, not implicit end-of-day dates.
    Use ``2024-12-31T23:00:00`` to include the last hourly observation of a year.
    """

    validation_fraction: float = 0.15
    calibration_fraction: float = 0.10
    test_fraction: float = 0.15
    gap_hours: int = 0
    train_end: str | pd.Timestamp | None = None
    validation_end: str | pd.Timestamp | None = None
    calibration_end: str | pd.Timestamp | None = None
    test_end: str | pd.Timestamp | None = None

    def __post_init__(self) -> None:
        fractions = (
            self.validation_fraction,
            self.calibration_fraction,
            self.test_fraction,
        )
        if any(isinstance(value, bool) or not isinstance(value, (int, float))
               or not math.isfinite(value) or value <= 0 for value in fractions) or sum(fractions) >= 1:
            raise ValueError(
                "validation_fraction + calibration_fraction + test_fraction "
                "must be positive and less than one"
            )
        if type(self.gap_hours) is not int or self.gap_hours < 0:
            raise ValueError("gap_hours must be a non-negative integer")
        present = [getattr(self, field) is not None for field in CALENDAR_FIELDS]
        if any(present) and not all(present):
            raise ValueError("Calendar splits require train_end, validation_end, calibration_end and test_end together")
        if all(present):
            dates = [_naive_timestamp(getattr(self, field), field) for field in CALENDAR_FIELDS]
            gap = pd.Timedelta(hours=self.gap_hours)
            if any(right <= left + gap for left, right in zip(dates, dates[1:])):
                raise ValueError("Calendar split end timestamps must increase and leave a non-empty interval after each purge gap")

    @classmethod
    def from_mapping(cls, values: Mapping[str, object]) -> "TemporalSplitConfig":
        """Read the shared fields used by model and experiment configurations."""
        return cls(
            validation_fraction=values.get("validation_fraction", 0.15),
            calibration_fraction=values.get("calibration_fraction", 0.10),
            test_fraction=values.get("test_fraction", 0.15),
            gap_hours=values.get("purge_gap_hours", values.get("gap_hours", 0)),
            **{field: values.get(field) for field in CALENDAR_FIELDS},
        )

    @property
    def split_mode(self) -> str:
        return "calendar" if self.train_end is not None else "fraction"

    @property
    def train_fraction(self) -> float:
        return 1.0 - self.validation_fraction - self.calibration_fraction - self.test_fraction


@dataclass(frozen=True)
class TemporalBoundaries:
    train_end: pd.Timestamp
    validation_end: pd.Timestamp
    calibration_end: pd.Timestamp
    gap_hours: int
    test_end: pd.Timestamp | None = None

    def to_dict(self) -> dict[str, object]:
        values = {
            "train_end": self.train_end.isoformat(),
            "validation_end": self.validation_end.isoformat(),
            "calibration_end": self.calibration_end.isoformat(),
            "gap_hours": self.gap_hours,
        }
        if self.test_end is not None:
            values.update(test_end=self.test_end.isoformat(), split_mode="calendar")
        return values

    @classmethod
    def from_dict(cls, values: Mapping[str, object]) -> "TemporalBoundaries":
        """Restore saved boundaries without recalculating them from new data."""
        train, validation, calibration = (
            _naive_timestamp(values[field], field) for field in CALENDAR_FIELDS[:3]
        )
        gap = values.get("gap_hours", 0)
        if type(gap) is not int or gap < 0:
            raise ValueError("gap_hours must be a non-negative integer")
        if not train < validation < calibration:
            raise ValueError("Stored split boundaries must be strictly increasing")
        test = _naive_timestamp(values["test_end"], "test_end") if values.get("test_end") is not None else None
        if values.get("split_mode") == "calendar" and test is None:
            raise ValueError("Stored calendar boundaries require test_end")
        if test is not None and test <= calibration + pd.Timedelta(hours=gap):
            raise ValueError("test_end must leave a non-empty Test interval after the purge gap")
        return cls(train, validation, calibration, gap, test)


def calendar_split_for_execution(values: Mapping[str, object], *, smoke: bool) -> tuple[dict, dict | None]:
    """Keep full-run dates; disclose fraction overrides for bounded smoke rows."""
    dates = {field: values.get(field) for field in CALENDAR_FIELDS}
    if smoke and any(value is not None for value in dates.values()):
        TemporalSplitConfig.from_mapping(values)  # Invalid dates are never hidden by smoke mode.
        return dict.fromkeys(CALENDAR_FIELDS), {
            "reason": "bounded_smoke_wiring_only",
            "requested_calendar": dates,
            "effective_split_mode": "fraction",
            "accuracy_evidence": False,
        }
    return dates, values.get("smoke_split_override") if smoke else None


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
        values = self._timestamps(timestamps).dropna().drop_duplicates().sort_values()
        if self.config.split_mode == "calendar":
            return TemporalBoundaries(
                **{field: _naive_timestamp(getattr(self.config, field), field) for field in CALENDAR_FIELDS},
                gap_hours=self.config.gap_hours,
            )
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
        values = self._timestamps(timestamps)
        gap = pd.Timedelta(hours=boundaries.gap_hours)
        labels = pd.Series(pd.NA, index=values.index, dtype="string")
        labels.loc[values.le(boundaries.train_end)] = "train"
        labels.loc[
            values.gt(boundaries.train_end + gap) & values.le(boundaries.validation_end)
        ] = "validation"
        labels.loc[
            values.gt(boundaries.validation_end + gap) & values.le(boundaries.calibration_end)
        ] = "calibration"
        test = values.gt(boundaries.calibration_end + gap)
        if boundaries.test_end is not None:
            test &= values.le(boundaries.test_end)
        labels.loc[test] = "test"
        return labels

    @staticmethod
    def _timestamps(timestamps: pd.Series | pd.Index) -> pd.Series:
        try:
            with warnings.catch_warnings():
                warnings.filterwarnings("error", message=".*mixed time zones.*", category=FutureWarning)
                values = pd.to_datetime(pd.Series(timestamps), errors="coerce", format="mixed")
        except (ValueError, TypeError, FutureWarning) as exc:
            raise ValueError("Split timestamps must use one timezone-naive local-time convention") from exc
        if not pd.api.types.is_datetime64_dtype(values.dtype):
            raise ValueError("Split timestamps must be timezone-naive (dataset local time, KST)")
        return values

    def split_frame(
        self,
        frame: pd.DataFrame,
        timestamp_column: str = "timestamp",
    ) -> TemporalFrameSplits:
        if timestamp_column not in frame:
            raise ValueError(f"Timestamp column is missing: {timestamp_column}")
        prepared = frame.copy()
        prepared[timestamp_column] = self._timestamps(prepared[timestamp_column])
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
