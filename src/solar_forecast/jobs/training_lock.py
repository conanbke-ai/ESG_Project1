"""동시 학습 실행을 차단하는 프로세스 잠금."""
from __future__ import annotations

from contextlib import contextmanager
import ctypes
import json
import os
from pathlib import Path
from typing import Iterator


class TrainingAlreadyRunning(RuntimeError):
    pass


def _pid_is_running(pid: int) -> bool:
    """Return whether the recorded owner PID still exists.

    Stale locks are removable only when a valid recorded PID is confirmed dead.
    Unknown/corrupt lock owners remain fail-closed.
    """

    if pid <= 0:
        return False
    if os.name == "nt":
        process_query_limited_information = 0x1000
        still_active = 259
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(
            process_query_limited_information,
            False,
            int(pid),
        )
        if not handle:
            return False
        try:
            code = ctypes.c_ulong()
            if not kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
                # Fail closed when Windows cannot determine the process state.
                return True
            return int(code.value) == still_active
        finally:
            kernel32.CloseHandle(handle)

    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    return True


def _lock_owner(lock_path: Path) -> tuple[dict | None, str]:
    try:
        raw = lock_path.read_text(encoding="utf-8", errors="replace")
    except FileNotFoundError:
        return None, "unknown"
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError:
        return None, raw
    if not isinstance(payload, dict):
        return None, raw
    try:
        pid = int(payload["pid"])
    except (KeyError, TypeError, ValueError):
        return None, raw
    payload["pid"] = pid
    return payload, raw


def _create_lock(lock_path: Path, model: str) -> int:
    descriptor = os.open(
        str(lock_path),
        os.O_CREAT | os.O_EXCL | os.O_WRONLY,
    )
    try:
        payload = json.dumps(
            {"pid": os.getpid(), "model": model},
            ensure_ascii=False,
        )
        os.write(descriptor, payload.encode("utf-8"))
    finally:
        os.close(descriptor)
    return descriptor


@contextmanager
def exclusive_training_lock(lock_path: Path, model: str) -> Iterator[None]:
    """Prevent concurrent training and recover locks whose owner process died."""

    lock_path.parent.mkdir(parents=True, exist_ok=True)

    for attempt in range(2):
        try:
            _create_lock(lock_path, model)
            break
        except FileExistsError as exc:
            owner, raw = _lock_owner(lock_path)
            if owner is None:
                raise TrainingAlreadyRunning(
                    f"Training lock exists with an unreadable owner: {raw}"
                ) from exc

            pid = int(owner["pid"])
            if _pid_is_running(pid):
                raise TrainingAlreadyRunning(
                    f"Another training job is active: {raw}"
                ) from exc

            # Confirmed stale owner. Remove once and retry the atomic create.
            try:
                lock_path.unlink()
            except FileNotFoundError:
                pass
            if attempt == 1:
                raise TrainingAlreadyRunning(
                    f"Could not recover stale training lock: {raw}"
                ) from exc
    else:
        raise TrainingAlreadyRunning("Could not acquire training lock")

    try:
        yield
    finally:
        try:
            lock_path.unlink()
        except FileNotFoundError:
            pass
