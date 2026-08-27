from __future__ import annotations

from contextlib import contextmanager
import json
import os
from pathlib import Path
from typing import Iterator


class TrainingAlreadyRunning(RuntimeError):
    pass


@contextmanager
def exclusive_training_lock(lock_path: Path, model: str) -> Iterator[None]:
    """Prevent XGBoost and CNN training from consuming resources concurrently."""
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        descriptor = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError as exc:
        owner = lock_path.read_text(encoding="utf-8", errors="replace") if lock_path.exists() else "unknown"
        raise TrainingAlreadyRunning(f"Another training job is active: {owner}") from exc
    try:
        payload = json.dumps({"pid": os.getpid(), "model": model}, ensure_ascii=False)
        os.write(descriptor, payload.encode("utf-8"))
        os.close(descriptor)
        yield
    finally:
        try:
            lock_path.unlink()
        except FileNotFoundError:
            pass
