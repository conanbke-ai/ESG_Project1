"""남동발전 홈페이지 화면 조작과 공식 첨부파일 다운로드."""
from __future__ import annotations

import calendar
import csv
from dataclasses import dataclass
from datetime import date, datetime, timedelta
import io
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import time

from solar_forecast.collectors.download_naming import build_solar_download_filename


@dataclass(frozen=True)
class KoenDownloadWindow:
    """Outcome for one inclusive provider query in KST."""

    start: date
    end: date
    status: str
    path: Path | None = None
    source_rows: int = 0
    detail: str = ""

    def evidence(self) -> dict[str, object]:
        return {"start_date": self.start.isoformat(), "end_date": self.end.isoformat(),
                "status": self.status, "source_rows": self.source_rows, "detail": self.detail,
                "path": str(self.path) if self.path else None}


class KoenBrowserDownloader:
    """Selenium adapter for KOEN's monthly generation download screen."""

    page_url = "https://www.koenergy.kr/kosep/gv/nf/dt/nfdt21/main.do?menuCd=FN0912020217"

    def __init__(
        self,
        output_root: Path,
        *,
        downloaded_on: date | None = None,
        overwrite: bool = False,
    ):
        self.output_root = output_root / "한국남동발전"
        self.downloaded_on = downloaded_on or date.today()
        self.overwrite = overwrite

    def download_year(
        self, year: int, *, start_month: int = 1, end_month: int = 12, max_retry: int = 3
    ) -> list[Path]:
        if not 1 <= start_month <= end_month <= 12:
            raise ValueError("months must satisfy 1 <= start_month <= end_month <= 12")
        start = date(year, start_month, 1)
        end = date(year, end_month, calendar.monthrange(year, end_month)[1])
        return [item.path for item in self.download_range(start, end, max_retry=max_retry) if item.path]

    def download_range(self, start: date, end: date, *, max_retry: int = 3) -> list[KoenDownloadWindow]:
        """Request the exact range in calendar-month chunks and report each outcome."""
        if end < start or max_retry < 1:
            raise ValueError("KOEN requires an ordered date range and at least one attempt")
        outcomes = []
        remaining = []
        for first, last in self._month_windows(start, end):
            destination = self._destination(first, last)
            if destination.is_file() and not self.overwrite:
                outcomes.append(self._inspect_window(first, last, destination, "cached"))
            else:
                remaining.append((first, last, destination))
        if not remaining:
            return outcomes

        self.output_root.mkdir(parents=True, exist_ok=True)
        with TemporaryDirectory(dir=self.output_root, prefix=".koen-download-") as temporary:
            download_dir = Path(temporary)
            driver = None
            try:
                driver = self._build_driver(download_dir)
                self._open(driver)
            except Exception as exc:
                outcomes.extend(KoenDownloadWindow(first, last, "browser_open_failed", detail=type(exc).__name__)
                                for first, last, _ in remaining)
            else:
                for first, last, destination in remaining:
                    outcomes.append(self._download_window(driver, download_dir, first, last, destination, max_retry))
            finally:
                if driver is not None:
                    driver.quit()
        return sorted(outcomes, key=lambda item: item.start)

    @staticmethod
    def _month_windows(start: date, end: date):
        while start <= end:
            month_end = start.replace(day=calendar.monthrange(start.year, start.month)[1])
            last = min(month_end, end)
            yield start, last
            if last == end:
                break
            start = last + timedelta(days=1)

    def _destination(self, start: date, end: date) -> Path:
        if start.day == 1 and end.day == calendar.monthrange(end.year, end.month)[1]:
            detail = f"월간통합_{start:%Y%m}"
        else:
            detail = f"기간통합_{start:%Y%m%d}_{end:%Y%m%d}"
        return self.output_root / str(start.year) / build_solar_download_filename(
            "한국남동발전(주)", detail, self.downloaded_on,
        )

    def _build_driver(self, download_dir: Path):
        from selenium import webdriver

        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", {
            "download.default_directory": str(download_dir.resolve()),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        })
        if os.getenv("KOEN_HEADLESS", "0") == "1":
            options.add_argument("--headless=new")
        return webdriver.Chrome(options=options)

    def _open(self, driver) -> None:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        driver.get(self.page_url)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "strDateS")))

    @staticmethod
    def submit_query(driver, start: date, end: date) -> None:
        """Submit the date form without invoking the separate CSV download action."""
        from selenium.webdriver.common.by import By

        driver.execute_script("document.getElementById('strDateS').value=arguments[0]", f"{start:%Y%m%d}")
        driver.execute_script("document.getElementById('strDateE').value=arguments[0]", f"{end:%Y%m%d}")
        driver.find_element(By.XPATH, "//a[contains(@href, 'goSubmit') or contains(text(),'조회')]").click()

    def _download_window(self, driver, download_dir, start, end, destination, max_retry) -> KoenDownloadWindow:
        from selenium.common.exceptions import UnexpectedAlertPresentException

        outcome = KoenDownloadWindow(start, end, "no_csv_downloaded")
        for _ in range(max_retry):
            try:
                self.submit_query(driver, start, end)
                time.sleep(3)
                before = set(download_dir.glob("*.csv"))
                driver.execute_script("goCsvDown();")
                downloaded = self._wait_for_csv(download_dir, before)
                if downloaded is None:
                    outcome = KoenDownloadWindow(start, end, "no_csv_downloaded")
                    continue
                outcome = self._inspect_window(start, end, downloaded, "downloaded")
                if outcome.path is None:
                    return outcome
                saved = self._move_original(downloaded, destination)
                return self._inspect_window(start, end, saved, "downloaded")
            except UnexpectedAlertPresentException as exc:
                return KoenDownloadWindow(start, end, "provider_alert", detail=" ".join((exc.alert_text or "").split())[:250])
            except Exception as exc:
                outcome = KoenDownloadWindow(start, end, "browser_request_failed", detail=type(exc).__name__)
        return outcome

    @staticmethod
    def _inspect_window(start: date, end: date, path: Path, status: str) -> KoenDownloadWindow:
        raw = path.read_bytes()
        for encoding in ("utf-8-sig", "cp949"):
            try:
                content = raw.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            return KoenDownloadWindow(start, end, "unexpected_csv", detail="unsupported encoding")
        try:
            rows = csv.reader(io.StringIO(content, newline=""), strict=True)
            header = [value.strip() for value in next(rows, [])]
            if not header:
                return KoenDownloadWindow(start, end, "empty_csv")
            if header.count("일자") != 1:
                return KoenDownloadWindow(start, end, "unexpected_csv", detail="missing date heading")
            date_index = header.index("일자")
            count = 0
            for row in rows:
                if not row:
                    continue
                # Legacy KOEN exports append one unlabeled, blank field to each data row.
                legacy_blank_field = len(row) == len(header) + 1 and not row[-1].strip()
                if len(row) != len(header) and not legacy_blank_field:
                    return KoenDownloadWindow(start, end, "unexpected_csv", detail="malformed row")
                value = row[date_index].strip().replace("/", "-")
                if value[4:5] == ".":
                    value = value[:10].replace(".", "-") + value[10:]
                day = datetime.fromisoformat(value).date()
                if not start <= day <= end:
                    return KoenDownloadWindow(start, end, "unexpected_csv", detail="row outside requested window or malformed")
                count += 1
        except (ValueError, csv.Error):
            return KoenDownloadWindow(start, end, "unexpected_csv", detail="invalid date or CSV syntax")
        if not count:
            return KoenDownloadWindow(start, end, "empty_csv")
        return KoenDownloadWindow(start, end, status, path, count)

    def _wait_for_csv(self, download_dir: Path, before: set[Path], timeout: int = 30) -> Path | None:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            candidates = [path for path in download_dir.glob("*.csv") if path not in before]
            if candidates and not list(download_dir.glob("*.crdownload")):
                return max(candidates, key=lambda path: path.stat().st_mtime)
            time.sleep(1)
        return None

    def _move_original(self, source: Path, destination: Path) -> Path:
        """Move provider bytes unchanged; encoding conversion belongs to Silver."""

        destination.parent.mkdir(parents=True, exist_ok=True)
        if destination.exists() and not self.overwrite:
            source.unlink(missing_ok=True)
            return destination
        source.replace(destination)
        return destination


def download_solar_data(
    base_path: str, year: int, max_retry: int = 3, start_month: int = 1, end_month: int = 12
):
    """Compatibility facade for the former root-level automation function."""
    return KoenBrowserDownloader(Path(base_path)).download_year(
        year, start_month=start_month, end_month=end_month, max_retry=max_retry
    )
