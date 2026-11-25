"""Git에서 제외된 .env.local을 읽어 미설정 환경변수에 적용."""
from __future__ import annotations

import os
from pathlib import Path


def load_local_env(path: Path | None = None) -> None:
    """Load missing environment values from a git-ignored local file."""
    env_path = path or Path(__file__).resolve().parents[3] / ".env.local"
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))
