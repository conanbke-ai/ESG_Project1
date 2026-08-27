from __future__ import annotations

from datetime import datetime, timezone
import json
import logging
from pathlib import Path
import traceback
from typing import Any, Mapping, Optional


logger = logging.getLogger(__name__)


def write_error_report(
    output_dir: Path,
    exc: BaseException,
    *,
    stage: str,
    context: Optional[Mapping[str, Any]] = None,
) -> Path:
    """Create an error file only after a failed run and return its path."""
    output_dir.mkdir(parents=True, exist_ok=True)
    error_path = output_dir / "error.log"
    safe_context = {key: str(value) for key, value in (context or {}).items()}
    header = {
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "stage": stage,
        "exception_type": type(exc).__name__,
        "message": str(exc),
        "context": safe_context,
    }
    body = json.dumps(header, ensure_ascii=False, indent=2) + "\n\nTRACEBACK\n" + "".join(
        traceback.format_exception(type(exc), exc, exc.__traceback__)
    )
    error_path.write_text(body, encoding="utf-8")
    logger.error("%s failed; error report: %s", stage, error_path)
    return error_path
