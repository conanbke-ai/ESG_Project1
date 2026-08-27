from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Optional, Sequence

import pandas as pd


SUPPORTED_EXTENSIONS = {".csv", ".csv.gz", ".xlsx", ".xls"}


def _data_extension(path: Path) -> str:
    return ".csv.gz" if path.name.lower().endswith(".csv.gz") else path.suffix.lower()


@dataclass(frozen=True)
class DatasetLoadPolicy:
    chunk_rows: int = 100_000
    memory_limit_mb: int = 1536
    numeric_dtype: str = "float32"

    def __post_init__(self) -> None:
        if self.chunk_rows <= 0 or self.memory_limit_mb <= 0:
            raise ValueError("chunk_rows and memory_limit_mb must be positive")
        if self.numeric_dtype not in {"float32", "float64"}:
            raise ValueError("numeric_dtype must be float32 or float64")


@dataclass(frozen=True)
class DatasetLoadReport:
    files: int
    input_bytes: int
    scanned_rows: int
    retained_rows: int
    selected_columns: tuple[str, ...]
    chunk_rows: int
    numeric_dtype: str
    dataframe_memory_mb: float
    memory_limit_mb: int

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


def discover_latest_file(input_dir: Path) -> Path:
    """Find the most recently modified supported dataset below ``input_dir``."""
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")
    candidates = [
        path
        for path in input_dir.rglob("*")
        if path.is_file() and _data_extension(path) in SUPPORTED_EXTENSIONS
    ]
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
        suffix = _data_extension(source)
        if suffix in {".csv", ".csv.gz"}:
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

    def load_training_frame(
        self,
        data_path: Path,
        *,
        columns: Sequence[str],
        numeric_columns: Sequence[str],
        equals_filters: dict[str, object] | None = None,
        truthy_filter: str | None = None,
        row_limit: int | None = None,
        policy: DatasetLoadPolicy | None = None,
    ) -> tuple[Path, pd.DataFrame, DatasetLoadReport]:
        """Read only model columns and push row filters into bounded CSV chunks."""

        source = Path(data_path)
        load_policy = policy or DatasetLoadPolicy()
        files = self._training_files(source)
        requested = list(dict.fromkeys(columns))
        filter_columns = list((equals_filters or {}).keys())
        if truthy_filter:
            filter_columns.append(truthy_filter)
        usecols = list(dict.fromkeys([*requested, *filter_columns]))
        parts: list[pd.DataFrame] = []
        scanned_rows = 0
        retained_rows = 0
        retained_bytes = 0
        for path in files:
            for chunk in pd.read_csv(
                path,
                usecols=usecols,
                chunksize=load_policy.chunk_rows,
                low_memory=False,
            ):
                scanned_rows += len(chunk)
                mask = pd.Series(True, index=chunk.index)
                for column, value in (equals_filters or {}).items():
                    mask &= chunk[column].astype(str).eq(str(value))
                if truthy_filter:
                    values = chunk[truthy_filter]
                    if not pd.api.types.is_bool_dtype(values):
                        values = values.astype(str).str.strip().str.lower().map(
                            {
                                "true": True,
                                "1": True,
                                "yes": True,
                                "false": False,
                                "0": False,
                                "no": False,
                            }
                        )
                    mask &= values.fillna(False)
                selected = chunk.loc[mask, requested].copy()
                for column in numeric_columns:
                    if column in selected:
                        selected[column] = pd.to_numeric(
                            selected[column], errors="coerce"
                        ).astype(load_policy.numeric_dtype)
                if row_limit is not None:
                    remaining = row_limit - retained_rows
                    selected = selected.head(max(0, remaining))
                if not selected.empty:
                    retained_rows += len(selected)
                    retained_bytes += int(selected.memory_usage(index=True, deep=True).sum())
                    if retained_bytes > load_policy.memory_limit_mb * 1024 * 1024:
                        raise MemoryError(
                            "Filtered training frame exceeded the configured memory budget: "
                            f"{retained_bytes / 1024 / 1024:.1f} MB > "
                            f"{load_policy.memory_limit_mb} MB. Reduce columns/period, use a larger "
                            "machine, or lower the admitted plant set; rows were not silently sampled."
                        )
                    parts.append(selected)
                if row_limit is not None and retained_rows >= row_limit:
                    break
            if row_limit is not None and retained_rows >= row_limit:
                break
        if not parts:
            raise ValueError("No rows remain after applying dataset loading filters")
        frame = pd.concat(parts, ignore_index=True, copy=False)
        for column in frame.select_dtypes(include="object"):
            if column != "timestamp" and frame[column].nunique(dropna=False) < max(1000, len(frame) // 10):
                frame[column] = frame[column].astype("category")
        actual_memory = int(frame.memory_usage(index=True, deep=True).sum())
        if actual_memory > load_policy.memory_limit_mb * 1024 * 1024:
            raise MemoryError(
                f"Concatenated training frame uses {actual_memory / 1024 / 1024:.1f} MB, above "
                f"the {load_policy.memory_limit_mb} MB budget"
            )
        report = DatasetLoadReport(
            files=len(files),
            input_bytes=sum(path.stat().st_size for path in files),
            scanned_rows=scanned_rows,
            retained_rows=len(frame),
            selected_columns=tuple(requested),
            chunk_rows=load_policy.chunk_rows,
            numeric_dtype=load_policy.numeric_dtype,
            dataframe_memory_mb=actual_memory / 1024 / 1024,
            memory_limit_mb=load_policy.memory_limit_mb,
        )
        return source, frame, report

    @staticmethod
    def _training_files(source: Path) -> list[Path]:
        if source.is_file():
            if _data_extension(source) not in {".csv", ".csv.gz"}:
                raise ValueError("Chunked training loading currently supports CSV and CSV.GZ")
            return [source]
        if source.is_dir():
            files = sorted(
                path
                for path in source.rglob("*")
                if path.is_file() and _data_extension(path) in {".csv", ".csv.gz"}
            )
            if files:
                return files
        raise FileNotFoundError(f"Training dataset or partition directory does not exist: {source}")


def load_dataset(data_path: Optional[Path], input_dir: Path) -> tuple[Path, pd.DataFrame]:
    """Compatibility facade around DatasetRepository."""
    return DatasetRepository(Path(input_dir)).load(data_path)
