"""수집기 인터페이스와 수집 결과 데이터 계약."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol


@dataclass(frozen=True)
class CollectionResult:
    source: str
    status: str
    files: list[Path] = field(default_factory=list)
    rows: int = 0
    message: str = ""


class Collector(Protocol):
    name: str

    def collect(self) -> CollectionResult: ...
