from __future__ import annotations

from pathlib import Path
from typing import Any, Optional, Protocol, Sequence

import pandas as pd

from .config import PipelineConfig
from .preprocessing import PreprocessResult


class DatasetPort(Protocol):
    def load(self, data_path: Optional[Path] = None) -> tuple[Path, pd.DataFrame]: ...


class PreprocessorPort(Protocol):
    def transform(
        self, frame: pd.DataFrame, target_column: str, feature_columns: Optional[Sequence[str]] = None
    ) -> PreprocessResult: ...


class TrainingPort(Protocol):
    def execute(
        self, frame: pd.DataFrame, features: list[str], config: PipelineConfig, run_dir: Path
    ) -> tuple[dict[str, Any], dict[str, Any]]: ...


class ReportPort(Protocol):
    def write(
        self, metrics: dict[str, Any], anomalies: pd.DataFrame, source: Path, output_path: Path
    ) -> Path: ...
