from __future__ import annotations

from pathlib import Path
from typing import Optional

import pandas as pd


SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls"}


def discover_latest_file(input_dir: Path) -> Path:
    """Find the most recently modified supported dataset below ``input_dir``."""
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")
    candidates = [p for p in input_dir.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS]
    if not candidates:
        raise FileNotFoundError(f"No CSV or Excel dataset found below: {input_dir}")
    return max(candidates, key=lambda path: path.stat().st_mtime)


class DatasetRepository:
    """Filesystem adapter for discovering and loading tabular model data."""

    def __init__(self, input_dir: Path):
        self.input_dir = input_dir

    def load(self, data_path: Optional[Path] = None) -> tuple[Path, pd.DataFrame]:
        source = Path(data_path) if data_path else discover_latest_file(self.input_dir)
        if not source.exists():
            raise FileNotFoundError(f"Dataset does not exist: {source}")
        suffix = source.suffix.lower()
        if suffix == ".csv":
            last_error: Optional[UnicodeDecodeError] = None
            for encoding in ("utf-8-sig", "utf-8", "cp949"):
                try:
                    return source, pd.read_csv(source, encoding=encoding)
                except UnicodeDecodeError as exc:
                    last_error = exc
            raise ValueError(f"Unable to decode CSV: {source}") from last_error
        if suffix in {".xlsx", ".xls"}:
            return source, pd.read_excel(source)
        raise ValueError(f"Unsupported dataset type: {suffix}")


def load_dataset(data_path: Optional[Path], input_dir: Path) -> tuple[Path, pd.DataFrame]:
    """Compatibility facade around DatasetRepository."""
    return DatasetRepository(Path(input_dir)).load(data_path)
