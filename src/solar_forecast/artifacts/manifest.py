from __future__ import annotations

from datetime import datetime, timezone
from hashlib import sha256
import json
from pathlib import Path
import time
from typing import Any


def sha256_file(path: Path, chunk_bytes: int = 1024 * 1024) -> str:
    """Hash an artifact incrementally so lineage checks stay memory bounded."""

    digest = sha256()
    with Path(path).open("rb") as stream:
        while chunk := stream.read(chunk_bytes):
            digest.update(chunk)
    return digest.hexdigest()


def replace_file_atomic(temporary: Path, destination: Path, attempts: int = 8) -> Path:
    """Atomically replace a file, retrying short Windows scanner/indexer locks."""

    temporary = Path(temporary)
    destination = Path(destination)
    for attempt in range(attempts):
        try:
            temporary.replace(destination)
            return destination
        except PermissionError:
            if attempt + 1 == attempts:
                raise
            time.sleep(min(0.05 * (2**attempt), 0.4))
    raise RuntimeError("unreachable atomic replace state")


def write_json_atomic(path: Path, payload: dict[str, Any]) -> Path:
    """Replace a JSON artifact only after its complete temporary write succeeds."""

    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return replace_file_atomic(temporary, path)


def write_manifest(path: Path, *, status: str, model: str, run_id: str, details: dict[str, Any]) -> Path:
    payload = {
        "status": status, "model": model, "run_id": run_id,
        "updated_at_utc": datetime.now(timezone.utc).isoformat(), "details": details,
    }
    return write_json_atomic(path, payload)
