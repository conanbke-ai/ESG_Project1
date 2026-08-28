from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re


_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\[\]\x00-\x1f]')
_CANONICAL_PATTERN = re.compile(
    r"^(?P<organization>.+)_\[(?P<detail>[^]]+)] 태양광발전실적_"
    r"(?P<download_date>\d{8})\.csv$"
)


def _clean_component(value: str) -> str:
    cleaned = _INVALID_FILENAME_CHARS.sub("_", str(value)).strip().strip(".")
    cleaned = re.sub(r"\s+", " ", cleaned)
    if not cleaned:
        raise ValueError("download filename components must not be empty")
    return cleaned


@dataclass(frozen=True)
class SolarDownloadName:
    """Canonical name for an immutable provider download (Bronze artifact)."""

    organization: str
    detail: str
    downloaded_on: date

    @property
    def filename(self) -> str:
        organization = _clean_component(self.organization)
        detail = _clean_component(self.detail)
        return (
            f"{organization}_[{detail}] 태양광발전실적_"
            f"{self.downloaded_on:%Y%m%d}.csv"
        )


def build_solar_download_filename(
    organization: str,
    detail: str,
    downloaded_on: date,
) -> str:
    return SolarDownloadName(organization, detail, downloaded_on).filename


def is_canonical_solar_download_name(path: str | Path) -> bool:
    return _CANONICAL_PATTERN.fullmatch(Path(path).name) is not None
