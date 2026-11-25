"""collector Silver 파일의 plant-hour 스키마 심사와 admission manifest."""
from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from solar_forecast.infrastructure.artifact_store import sha256_file, write_json_atomic

from solar_forecast.collectors.plant_identity import resolve_generation_identity
from solar_forecast.collectors.generation_normalizers import (
    GENERATION_COLUMNS,
    GENERATION_CONTRACT_VERSION,
    read_csv_with_fallback,
)


@dataclass(frozen=True)
class CollectedGenerationFileAdmission:
    source: str
    status: str
    reason: str | None
    rows: int
    plants: int
    start: str | None
    end: str | None
    source_bytes: int
    source_sha256: str
    identity_resolutions: tuple[dict[str, object], ...] = ()


@dataclass(frozen=True)
class CollectedGenerationAdmissionResult:
    source_dir: Path
    manifest_path: Path
    files: tuple[CollectedGenerationFileAdmission, ...]

    @property
    def accepted_paths(self) -> tuple[Path, ...]:
        return tuple(
            Path(item.source) for item in self.files if item.status == "accepted"
        )

    @property
    def accepted_count(self) -> int:
        return len(self.accepted_paths)

    @property
    def rejected_count(self) -> int:
        return sum(item.status == "rejected" for item in self.files)

    @property
    def accepted_rows(self) -> int:
        return sum(item.rows for item in self.files if item.status == "accepted")


class CollectedGenerationAdmissionService:
    """Promote collector Silver files only when they satisfy the Gold contract."""

    required_columns = set(GENERATION_COLUMNS)

    def __init__(self, source_dir: Path, manifest_path: Path):
        self.source_dir = Path(source_dir)
        self.manifest_path = Path(manifest_path)

    def run(self) -> CollectedGenerationAdmissionResult:
        files = tuple(sorted(self.source_dir.rglob("*.csv"))) if self.source_dir.exists() else ()
        admissions = tuple(self._inspect(path) for path in files)
        payload = {
            "created_at": datetime.now().isoformat(),
            "contract": "collector-generation-admission.v1",
            "generation_contract_version": GENERATION_CONTRACT_VERSION,
            "source_dir": str(self.source_dir),
            "policy": (
                "Collector Silver files are admitted to Gold only when they expose "
                "the canonical plant-hour generation schema. Region-level training "
                "exports and files without plant identity remain archived but are not "
                "fed into the plant registry."
            ),
            "summary": {
                "files": len(admissions),
                "accepted_files": sum(item.status == "accepted" for item in admissions),
                "rejected_files": sum(item.status == "rejected" for item in admissions),
                "accepted_rows": sum(
                    item.rows for item in admissions if item.status == "accepted"
                ),
            },
            "files": [asdict(item) for item in admissions],
        }
        write_json_atomic(self.manifest_path, payload)
        return CollectedGenerationAdmissionResult(
            self.source_dir,
            self.manifest_path,
            admissions,
        )

    def _inspect(self, path: Path) -> CollectedGenerationFileAdmission:
        source_bytes = path.stat().st_size
        source_sha256 = sha256_file(path)
        try:
            header = read_csv_with_fallback(path, index_col=False, nrows=0)
            columns = {str(column).strip() for column in header.columns}
            missing = sorted(self.required_columns - columns)
            if missing:
                return CollectedGenerationFileAdmission(
                    source=str(path),
                    status="rejected",
                    reason="missing_generation_contract_columns: " + ", ".join(missing),
                    rows=0,
                    plants=0,
                    start=None,
                    end=None,
                    source_bytes=source_bytes,
                    source_sha256=source_sha256,
                )
            frame = self._read_generation_columns(path)
            if frame.empty:
                return CollectedGenerationFileAdmission(
                    source=str(path),
                    status="rejected",
                    reason="no_valid_generation_rows",
                    rows=0,
                    plants=0,
                    start=None,
                    end=None,
                    source_bytes=source_bytes,
                    source_sha256=source_sha256,
                )
            return CollectedGenerationFileAdmission(
                source=str(path),
                status="accepted",
                reason=None,
                rows=len(frame),
                plants=int(frame["plant_id"].nunique()),
                start=frame["timestamp"].min().isoformat(),
                end=frame["timestamp"].max().isoformat(),
                source_bytes=source_bytes,
                source_sha256=source_sha256,
                identity_resolutions=tuple(frame.attrs.get("plant_identity_resolutions", ())),
            )
        except Exception as exc:
            return CollectedGenerationFileAdmission(
                source=str(path),
                status="rejected",
                reason=f"read_error: {exc}",
                rows=0,
                plants=0,
                start=None,
                end=None,
                source_bytes=source_bytes,
                source_sha256=source_sha256,
            )

    @staticmethod
    def _read_generation_columns(path: Path) -> pd.DataFrame:
        frame = read_csv_with_fallback(path, index_col=False)
        frame.columns = [str(column).strip() for column in frame.columns]
        frame = frame[GENERATION_COLUMNS].copy()
        frame["timestamp"] = pd.to_datetime(frame["timestamp"], errors="coerce")
        frame["generation_mwh"] = pd.to_numeric(
            frame["generation_mwh"], errors="coerce"
        )
        return resolve_generation_identity(
            frame.dropna(subset=["timestamp", "generation_mwh"])
        )


def collect_generation_admissions(
    source_dir: Path,
    manifest_path: Path,
) -> CollectedGenerationAdmissionResult:
    return CollectedGenerationAdmissionService(source_dir, manifest_path).run()


__all__ = [
    "CollectedGenerationAdmissionResult",
    "CollectedGenerationAdmissionService",
    "CollectedGenerationFileAdmission",
    "collect_generation_admissions",
]
