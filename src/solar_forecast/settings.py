from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[2]


@dataclass(frozen=True)
class ModelJobConfig:
    model: str
    profile: str
    values: dict[str, Any]
    source: Path


def load_model_config(path: Path) -> ModelJobConfig:
    resolved = path if path.is_absolute() else PROJECT_ROOT / path
    values = json.loads(resolved.read_text(encoding="utf-8"))
    if values.get("model") not in {"xgboost", "cnn_bilstm", "hybrid"}:
        raise ValueError(f"Unsupported model in config: {values.get('model')}")
    return ModelJobConfig(values["model"], values.get("profile", "optimized"), values, resolved)
