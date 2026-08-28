from __future__ import annotations

import calendar
from datetime import date
import os
from pathlib import Path
import time

from .naming import build_solar_download_filename


class KoenBrowserDownloader:
    """Selenium adapter for KOEN's monthly generation download screen."""

    page_url = "https://www.koenergy.kr/kosep/gv/nf/dt/nfdt21/main.do"

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
        self.output_root.mkdir(parents=True, exist_ok=True)
        year_dir = self.output_root / str(year)
        year_dir.mkdir(parents=True, exist_ok=True)
        driver = self._build_driver()
        outputs: list[Path] = []
        try:
            self._open(driver)
            for month in range(start_month, end_month + 1):
                output = self._download_month(driver, year_dir, year, month, max_retry)
                if output is not None:
                    outputs.append(output)
        finally:
            driver.quit()
        return outputs

    def _build_driver(self):
        from selenium import webdriver

        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", {
            "download.default_directory": str(self.output_root.resolve()),
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

    def _download_month(self, driver, year_dir: Path, year: int, month: int, max_retry: int) -> Path | None:
        from selenium.webdriver.common.by import By

        start_date = f"{year}{month:02d}01"
        end_date = f"{year}{month:02d}{calendar.monthrange(year, month)[1]:02d}"
        destination = year_dir / build_solar_download_filename(
            "한국남동발전(주)",
            f"월간통합_{year}{month:02d}",
            self.downloaded_on,
        )
        for _ in range(max_retry):
            try:
                driver.execute_script("document.getElementById('strDateS').value=arguments[0]", start_date)
                driver.execute_script("document.getElementById('strDateE').value=arguments[0]", end_date)
                driver.find_element(By.XPATH, "//a[contains(@href, 'goSubmit') or contains(text(),'조회')]").click()
                time.sleep(3)
                before = set(self.output_root.glob("*.csv"))
                driver.execute_script("goCsvDown();")
                downloaded = self._wait_for_csv(before)
                if downloaded is not None:
                    return self._move_original(downloaded, destination)
            except Exception:
                time.sleep(2)
        return None

    def _wait_for_csv(self, before: set[Path], timeout: int = 30) -> Path | None:
        deadline = time.time() + timeout
        while time.time() < deadline:
            candidates = [path for path in self.output_root.glob("*.csv") if path not in before]
            if candidates and not list(self.output_root.glob("*.crdownload")):
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
