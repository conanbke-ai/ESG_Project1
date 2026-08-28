from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

import requests

from .base import CollectionResult
from .config import CollectionConfig
from .naming import build_solar_download_filename


class FileNormalizer(Protocol):
    def write(self, source: Path, destination: Path) -> Path: ...


def _csv_data_rows(path: Path) -> int:
    with path.open("rb") as source:
        return max(sum(1 for _ in source) - 1, 0)


@dataclass(frozen=True)
class EwpAttachmentSpec:
    detail_url: str
    download_url: str
    attachment_id: str
    order_number: str
    page_code: str
    organization: str
    detail_name: str


class EwpTrainingDataCollector:
    """Download EWP's nationwide solar training CSV from its public-data board."""

    name = "ewp"
    def __init__(self, config: CollectionConfig, spec: EwpAttachmentSpec):
        self.config = config
        self.spec = spec

    def collect(self) -> CollectionResult:
        filename = build_solar_download_filename(
            self.spec.organization,
            self.spec.detail_name,
            self.config.download_date,
        )
        destination = self.config.output_dir / self.name / filename
        destination.parent.mkdir(parents=True, exist_ok=True)
        status = "cached"
        if self.config.overwrite or not destination.exists():
            temporary = destination.with_suffix(".csv.part")
            with requests.Session() as session:
                detail = session.get(self.spec.detail_url, timeout=30)
                detail.raise_for_status()
                response = session.post(
                    self.spec.download_url,
                    data={
                        "idx": self.spec.attachment_id,
                        "page": "1",
                        "goPage": "/kor/subpage/content",
                        "pc": self.spec.page_code,
                        "idx_to": self.spec.attachment_id,
                        "order_num": self.spec.order_number,
                    },
                    headers={"Referer": self.spec.detail_url},
                    timeout=120,
                    stream=True,
                )
                response.raise_for_status()
                content_type = response.headers.get("Content-Type", "").lower()
                if "text/html" in content_type:
                    raise ValueError("EWP download endpoint returned HTML instead of the CSV attachment")
                with temporary.open("wb") as output:
                    for chunk in response.iter_content(chunk_size=1024 * 1024):
                        if chunk:
                            output.write(chunk)
            temporary.replace(destination)
            status = "downloaded"
        from .normalization import EwpTrainingNormalizer

        normalized = self.config.standardized_output_dir / self.name / f"{destination.stem}_표준화.csv"
        EwpTrainingNormalizer().write(destination, normalized)
        return CollectionResult(
            self.name,
            status,
            [destination, normalized],
            rows=_csv_data_rows(normalized),
            message="EWP nationwide solar training data refreshed from the official public-data board",
        )


@dataclass(frozen=True)
class DataGoDatasetSpec:
    source: str
    dataset_id: str
    detail_id: str
    organization: str
    detail_name: str

    @property
    def detail_url(self) -> str:
        return f"https://www.data.go.kr/data/{self.dataset_id}/fileData.do"


class DataGoFileCollector:
    """Download a login-free official attachment through data.go.kr's public flow."""

    metadata_url = "https://www.data.go.kr/tcs/dss/selectFileDataDownload.do"
    limit_url = "https://www.data.go.kr/cmm/cmm/check-limit.json"
    download_url = "https://www.data.go.kr/cmm/cmm/fileDownload.do"

    def __init__(
        self,
        spec: DataGoDatasetSpec,
        config: CollectionConfig,
        normalizer: FileNormalizer | None = None,
    ):
        self.spec = spec
        self.config = config
        self.normalizer = normalizer

    def collect(self) -> CollectionResult:
        filename = build_solar_download_filename(
            self.spec.organization,
            self.spec.detail_name,
            self.config.download_date,
        )
        destination = self.config.output_dir / self.spec.source / filename
        destination.parent.mkdir(parents=True, exist_ok=True)
        status = "cached"
        if self.config.overwrite or not destination.exists():
            temporary = destination.with_suffix(destination.suffix + ".part")
            with requests.Session() as session:
                detail = session.get(self.spec.detail_url, timeout=30)
                detail.raise_for_status()
                metadata_response = session.get(
                    self.metadata_url,
                    params={
                        "recommendDataYn": "Y",
                        "publicDataPk": self.spec.dataset_id,
                        "publicDataDetailPk": self.spec.detail_id,
                    },
                    headers={"Referer": self.spec.detail_url},
                    timeout=30,
                )
                metadata_response.raise_for_status()
                metadata = metadata_response.json()
                if not metadata.get("status"):
                    raise ValueError(f"data.go.kr did not return attachment metadata for {self.spec.dataset_id}")
                attachment_id = str(metadata["atchFileId"])
                detail_sn = str(metadata["fileDetailSn"])
                data_name = str(metadata["dataSetFileDetailInfo"]["dataNm"])
                limit = session.post(
                    self.limit_url,
                    data={"atchFileId": attachment_id, "fileDetailSn": detail_sn},
                    headers={"Referer": self.spec.detail_url},
                    timeout=30,
                )
                limit.raise_for_status()
                if limit.json().get("needCaptcha"):
                    raise RuntimeError("data.go.kr requires a CAPTCHA for this download; retry after the rate limit resets")
                response = session.get(
                    self.download_url,
                    params={"atchFileId": attachment_id, "fileDetailSn": detail_sn, "dataNm": data_name},
                    headers={"Referer": self.spec.detail_url},
                    timeout=120,
                    stream=True,
                )
                response.raise_for_status()
                if "text/html" in response.headers.get("Content-Type", "").lower():
                    raise ValueError("data.go.kr returned HTML instead of the requested attachment")
                with temporary.open("wb") as output:
                    for chunk in response.iter_content(chunk_size=1024 * 1024):
                        if chunk:
                            output.write(chunk)
            temporary.replace(destination)
            status = "downloaded"
        files = [destination]
        if self.normalizer is not None:
            normalized = (
                self.config.standardized_output_dir
                / self.spec.source
                / f"{destination.stem}_표준화.csv"
            )
            self.normalizer.write(destination, normalized)
            files.append(normalized)
        return CollectionResult(
            self.spec.source,
            status,
            files,
            rows=_csv_data_rows(files[-1]),
            message=f"Official data.go.kr dataset {self.spec.dataset_id} refreshed",
        )


class KoenHomepageCollector:
    """Reuse the existing official KOEN monthly-download browser automation."""

    name = "koen"

    def __init__(self, config: CollectionConfig):
        self.config = config

    def collect(self) -> CollectionResult:
        from .koen_browser import KoenBrowserDownloader
        from .normalization import KoenGenerationNormalizer

        years = range(self.config.start_date.year, self.config.end_date.year + 1)
        base = self.config.output_dir / "koen"
        downloader = KoenBrowserDownloader(
            base,
            downloaded_on=self.config.download_date,
            overwrite=self.config.overwrite,
        )
        files: list[Path] = []
        for year in years:
            start_month = self.config.start_date.month if year == self.config.start_date.year else 1
            end_month = self.config.end_date.month if year == self.config.end_date.year else 12
            files.extend(downloader.download_year(year, start_month=start_month, end_month=end_month))
        normalized_files: list[Path] = []
        normalizer = KoenGenerationNormalizer()
        for source in files:
            destination = (
                self.config.standardized_output_dir
                / self.name
                / f"{source.stem}_표준화.csv"
            )
            normalized_files.append(normalizer.write(source, destination))
        return CollectionResult(
            self.name,
            "downloaded",
            [*files, *normalized_files],
            rows=sum(_csv_data_rows(path) for path in normalized_files),
        )
