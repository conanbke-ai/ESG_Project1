from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from hashlib import sha256
import json
from pathlib import Path
import random
import re
from typing import Any, Mapping

import numpy as np
import pandas as pd
import torch

from solar_forecast.artifacts.manifest import (
    replace_file_atomic,
    sha256_file,
    write_json_atomic,
)
from solar_forecast.settings import ModelJobConfig, PROJECT_ROOT


CHECKPOINT_CONTRACT = "solar-training-checkpoint.v2"


def stable_signature(payload: Mapping[str, Any] | list[Any] | tuple[Any, ...]) -> str:
    encoded = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        default=str,
    ).encode("utf-8")
    return sha256(encoded).hexdigest()


def dataframe_signature(
    frame: pd.DataFrame,
    columns: list[str],
) -> str:
    """Hash selected frame values without serializing a second tabular copy."""

    columns = list(dict.fromkeys(columns))
    missing = [column for column in columns if column not in frame]
    if missing:
        raise ValueError(f"Cannot fingerprint missing frame columns: {missing}")
    digest = sha256()
    digest.update(stable_signature({"columns": columns, "rows": len(frame)}).encode("ascii"))
    hashes = pd.util.hash_pandas_object(
        frame[columns],
        index=True,
        categorize=True,
    ).to_numpy(dtype=np.uint64, copy=False)
    digest.update(hashes.tobytes())
    return digest.hexdigest()


def dataset_signature(source: Path) -> str:
    """Fingerprint one file or partitioned dataset independently of the model."""

    source = Path(source)
    if source.is_file():
        return sha256_file(source)
    manifest = source.parent / "model_ready_manifest.json"
    inventory = [
        {
            "path": path.relative_to(source).as_posix(),
            "bytes": path.stat().st_size,
            "mtime_ns": path.stat().st_mtime_ns,
        }
        for path in sorted(source.rglob("*"))
        if path.is_file()
    ]
    return stable_signature(
        {
            "manifest_sha256": sha256_file(manifest) if manifest.exists() else None,
            "partition_inventory": inventory,
        }
    )


def training_fingerprint(config: ModelJobConfig) -> str:
    values = json.loads(json.dumps(config.values, ensure_ascii=False, default=str))
    values.pop("resume", None)
    values.pop("checkpoint", None)
    optimizer = values.get("optimizer")
    if isinstance(optimizer, dict):
        for operational_key in (
            "max_trials",
            "timeout_seconds",
            "heartbeat_interval_seconds",
            "grace_period_seconds",
            "max_failed_trial_retries",
        ):
            optimizer.pop(operational_key, None)
    source = Path(str(config.values.get("input_dataset", "")))
    source = source if source.is_absolute() else PROJECT_ROOT / source
    return stable_signature(
        {
            "checkpoint_contract": CHECKPOINT_CONTRACT,
            "model": config.model,
            "profile": config.profile,
            "model_values": values,
            "dataset": dataset_signature(source),
        }
    )


def capture_rng_state() -> dict[str, Any]:
    return {
        "python": random.getstate(),
        "numpy": np.random.get_state(),
        "torch_cpu": torch.get_rng_state(),
        "torch_cuda": torch.cuda.get_rng_state_all() if torch.cuda.is_available() else None,
    }


def restore_rng_state(state: Mapping[str, Any] | None) -> None:
    if not state:
        return
    random.setstate(state["python"])
    np.random.set_state(state["numpy"])
    torch.set_rng_state(state["torch_cpu"])
    cuda_state = state.get("torch_cuda")
    if cuda_state is not None and torch.cuda.is_available():
        torch.cuda.set_rng_state_all(cuda_state)


@dataclass(frozen=True)
class CheckpointSettings:
    enabled: bool = True
    resume: bool = True
    root: Path = Path("artifacts/checkpoints")
    cnn_every_epochs: int = 1
    xgboost_every_rounds: int = 50

    @classmethod
    def from_config(cls, config: ModelJobConfig) -> "CheckpointSettings":
        raw = config.values.get("checkpoint", {})
        if not isinstance(raw, Mapping):
            raise ValueError("checkpoint configuration must be an object")
        settings = cls(
            enabled=bool(raw.get("enabled", True)),
            resume=bool(raw.get("resume", config.values.get("resume", True))),
            root=Path(str(raw.get("root", "artifacts/checkpoints"))),
            cnn_every_epochs=int(raw.get("cnn_every_epochs", 1)),
            xgboost_every_rounds=int(raw.get("xgboost_every_rounds", 50)),
        )
        if min(settings.cnn_every_epochs, settings.xgboost_every_rounds) < 1:
            raise ValueError("checkpoint save intervals must be positive")
        return settings


class TrainingCheckpointStore:
    """Atomic, fingerprint-scoped state for interrupted model training."""

    def __init__(
        self,
        root: Path,
        *,
        model: str,
        fingerprint: str,
        enabled: bool = True,
        resume: bool = True,
        cnn_every_epochs: int = 1,
        xgboost_every_rounds: int = 50,
    ):
        self.root = Path(root)
        self.model = model
        self.fingerprint = fingerprint
        self.enabled = enabled
        self.resume = resume
        self.cnn_every_epochs = cnn_every_epochs
        self.xgboost_every_rounds = xgboost_every_rounds

    @classmethod
    def from_config(
        cls,
        config: ModelJobConfig,
        *,
        project_root: Path = PROJECT_ROOT,
    ) -> "TrainingCheckpointStore":
        settings = CheckpointSettings.from_config(config)
        root = settings.root if settings.root.is_absolute() else Path(project_root) / settings.root
        return cls(
            root,
            model=config.model,
            fingerprint=training_fingerprint(config),
            enabled=settings.enabled,
            resume=settings.resume,
            cnn_every_epochs=settings.cnn_every_epochs,
            xgboost_every_rounds=settings.xgboost_every_rounds,
        )

    @property
    def directory(self) -> Path:
        return self.root / self.model / self.fingerprint

    def torch_path(self, stage: str) -> Path:
        return self.directory / f"{self._stage_name(stage)}.pt"

    def xgboost_path(self, stage: str) -> Path:
        return self.directory / f"{self._stage_name(stage)}.json"

    def save_torch(
        self,
        stage: str,
        payload: Mapping[str, Any],
        *,
        signature: str,
        progress: Mapping[str, Any],
        completed: bool = False,
    ) -> Path | None:
        if not self.enabled:
            return None
        path = self.torch_path(stage)
        path.parent.mkdir(parents=True, exist_ok=True)
        state = {
            **dict(payload),
            "checkpoint_contract": CHECKPOINT_CONTRACT,
            "model": self.model,
            "fingerprint": self.fingerprint,
            "stage": stage,
            "signature": signature,
            "completed": completed,
            "saved_at_utc": datetime.now(timezone.utc).isoformat(),
        }
        temporary = path.with_name(f"{path.stem}.tmp{path.suffix}")
        torch.save(state, temporary)
        replace_file_atomic(temporary, path)
        self._write_metadata(path, stage, signature, progress, completed)
        return path

    def load_torch(
        self,
        stage: str,
        *,
        signature: str,
        map_location: torch.device | str = "cpu",
    ) -> dict[str, Any] | None:
        if not self.enabled or not self.resume:
            return None
        path = self.torch_path(stage)
        if not path.exists():
            return None
        state = torch.load(path, map_location=map_location, weights_only=False)
        self._validate_state(state, stage, signature)
        return dict(state)

    def save_xgboost(
        self,
        booster: Any,
        stage: str,
        *,
        signature: str,
        completed_rounds: int,
        completed: bool = False,
    ) -> Path | None:
        if not self.enabled:
            return None
        path = self.xgboost_path(stage)
        path.parent.mkdir(parents=True, exist_ok=True)
        booster.set_attr(
            solar_checkpoint_contract=CHECKPOINT_CONTRACT,
            solar_checkpoint_model=self.model,
            solar_checkpoint_fingerprint=self.fingerprint,
            solar_checkpoint_stage=stage,
            solar_checkpoint_signature=signature,
            solar_checkpoint_completed="true" if completed else "false",
            solar_checkpoint_rounds=str(int(completed_rounds)),
        )
        temporary = path.with_name(f"{path.stem}.tmp{path.suffix}")
        booster.save_model(temporary)
        replace_file_atomic(temporary, path)
        self._write_metadata(
            path,
            stage,
            signature,
            {"completed_rounds": int(completed_rounds)},
            completed,
        )
        return path

    def load_xgboost(
        self,
        stage: str,
        *,
        signature: str,
    ) -> tuple[Path, dict[str, Any]] | None:
        if not self.enabled or not self.resume:
            return None
        path = self.xgboost_path(stage)
        metadata_path = self._metadata_path(path)
        if not path.exists():
            return None
        metadata = (
            json.loads(metadata_path.read_text(encoding="utf-8"))
            if metadata_path.exists()
            else None
        )
        metadata_error: ValueError | None = None
        if metadata is not None:
            try:
                self._validate_state(metadata, stage, signature)
            except ValueError as exc:
                metadata_error = exc

        # The model embeds the same contract, allowing safe recovery when a
        # process stops after replacing the Booster but before metadata replace.
        import xgboost as xgb

        booster = xgb.Booster()
        booster.load_model(path)
        embedded = {
            "checkpoint_contract": booster.attr("solar_checkpoint_contract"),
            "model": booster.attr("solar_checkpoint_model"),
            "fingerprint": booster.attr("solar_checkpoint_fingerprint"),
            "stage": booster.attr("solar_checkpoint_stage"),
            "signature": booster.attr("solar_checkpoint_signature"),
        }
        if all(embedded.values()):
            self._validate_state(embedded, stage, signature)
        elif metadata is None or metadata_error is not None:
            if metadata_error is not None:
                raise metadata_error
            return None

        actual_rounds = int(booster.num_boosted_rounds())
        if metadata is None or metadata_error is not None:
            metadata = {
                **embedded,
                "completed": booster.attr("solar_checkpoint_completed") == "true",
                "checkpoint_path": str(path),
            }
        elif booster.attr("solar_checkpoint_completed") is not None:
            metadata["completed"] = (
                booster.attr("solar_checkpoint_completed") == "true"
            )
        metadata["progress"] = {"completed_rounds": actual_rounds}
        return path, metadata

    def remove(self, stage: str, *, kind: str) -> None:
        if not self.enabled:
            return
        if kind == "torch":
            path = self.torch_path(stage)
        elif kind == "xgboost":
            path = self.xgboost_path(stage)
        else:
            raise ValueError("checkpoint kind must be torch or xgboost")
        for target in (path, self._metadata_path(path)):
            if target.exists():
                target.unlink()

    def describe(self) -> dict[str, Any]:
        return {
            "enabled": self.enabled,
            "resume": self.resume,
            "contract": CHECKPOINT_CONTRACT,
            "fingerprint": self.fingerprint,
            "directory": str(self.directory),
            "cnn_every_epochs": self.cnn_every_epochs,
            "xgboost_every_rounds": self.xgboost_every_rounds,
        }

    def _write_metadata(
        self,
        path: Path,
        stage: str,
        signature: str,
        progress: Mapping[str, Any],
        completed: bool,
    ) -> None:
        write_json_atomic(
            self._metadata_path(path),
            {
                "checkpoint_contract": CHECKPOINT_CONTRACT,
                "model": self.model,
                "fingerprint": self.fingerprint,
                "stage": stage,
                "signature": signature,
                "completed": completed,
                "saved_at_utc": datetime.now(timezone.utc).isoformat(),
                "checkpoint_path": str(path),
                "progress": dict(progress),
            },
        )

    def _validate_state(
        self,
        state: Mapping[str, Any],
        stage: str,
        signature: str,
    ) -> None:
        expected = {
            "checkpoint_contract": CHECKPOINT_CONTRACT,
            "model": self.model,
            "fingerprint": self.fingerprint,
            "stage": stage,
            "signature": signature,
        }
        mismatched = [key for key, value in expected.items() if state.get(key) != value]
        if mismatched:
            raise ValueError(
                "Checkpoint is incompatible with the current data/configuration: "
                f"{mismatched}"
            )

    @staticmethod
    def _metadata_path(path: Path) -> Path:
        return path.with_suffix(path.suffix + ".meta.json")

    @staticmethod
    def _stage_name(stage: str) -> str:
        normalized = re.sub(r"[^A-Za-z0-9_.-]+", "_", stage).strip("._")
        if not normalized:
            raise ValueError("checkpoint stage cannot be blank")
        return normalized
