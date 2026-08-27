from __future__ import annotations

from datetime import datetime
import importlib
from pathlib import Path
from typing import Callable

from solar_forecast.artifacts.manifest import write_manifest
from solar_forecast.jobs.lock import exclusive_training_lock
from solar_forecast.settings import ModelJobConfig, PROJECT_ROOT


TRAINERS = {
    "xgboost": ("solar_forecast.models.xgboost", "train"),
    "cnn_bilstm": ("solar_forecast.models.cnn_bilstm", "train"),
}


class TrainingService:
    """Runs one model strategy behind a process-wide lock and manifest boundary."""

    def __init__(self, trainers: dict[str, tuple[str, str]] | None = None):
        self.trainers = trainers or TRAINERS

    def run(self, config: ModelJobConfig, *, smoke: bool = False) -> Path:
        if config.model not in self.trainers:
            raise ValueError("Hybrid does not train base models; use the hybrid command")
        run_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        run_dir = PROJECT_ROOT / config.values["output_root"] / run_id
        manifest_path = run_dir / "manifest.json"
        lock_path = PROJECT_ROOT / "artifacts" / ".training.lock"
        write_manifest(manifest_path, status="running", model=config.model, run_id=run_id, details={"smoke": smoke})
        try:
            with exclusive_training_lock(lock_path, config.model):
                module_name, function_name = self.trainers[config.model]
                trainer: Callable = getattr(importlib.import_module(module_name), function_name)
                result = trainer(config, run_dir=run_dir, smoke=smoke)
            write_manifest(manifest_path, status="completed", model=config.model, run_id=run_id, details=result)
            return run_dir
        except Exception as exc:
            write_manifest(manifest_path, status="failed", model=config.model, run_id=run_id, details={"error": str(exc)})
            raise


def run_training(config: ModelJobConfig, *, smoke: bool = False) -> Path:
    """Compatibility facade around TrainingService."""
    return TrainingService().run(config, smoke=smoke)
