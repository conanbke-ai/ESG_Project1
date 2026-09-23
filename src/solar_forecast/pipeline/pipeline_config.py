"""전체 예측 파이프라인의 입력·학습·보고서 실행 설정."""
from __future__ import annotations

from solar_forecast.infrastructure.project_paths import LEGACY_MERGED_ROOT, PIPELINE_RUNS_ROOT

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Sequence

from solar_forecast.models.cnn_bilstm.sequence_config import SequenceConfig


@dataclass(frozen=True)
class PipelineConfig:
    """Configuration shared by every pipeline stage."""

    target_column: str
    data_path: Optional[Path] = None
    input_dir: Path = LEGACY_MERGED_ROOT
    feature_columns: Optional[Sequence[str]] = None
    output_dir: Path = PIPELINE_RUNS_ROOT
    sequence: SequenceConfig = field(default_factory=SequenceConfig)
    epochs: int = 50
    n_trials: int = 10
    use_optuna: bool = True
    optimizer_timeout_seconds: int | None = None
    use_reinforcement: bool = False
    contamination: float = 0.05
    artifact_level: str = "minimal"

    def __post_init__(self) -> None:
        if self.artifact_level not in {"minimal", "standard", "debug"}:
            raise ValueError("artifact_level must be one of: minimal, standard, debug")
        if self.n_trials < 1:
            raise ValueError("n_trials must be positive")
        if self.optimizer_timeout_seconds is not None and self.optimizer_timeout_seconds < 1:
            raise ValueError("optimizer_timeout_seconds must be positive or null")
