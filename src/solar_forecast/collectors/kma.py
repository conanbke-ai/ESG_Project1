from __future__ import annotations

import csv
from datetime import date, timedelta
import os
from pathlib import Path
import time

from .base import CollectionResult
from .config import CollectionConfig


class KmaAsosHourlyCollector:
    """Incrementally download ASOS hourly CSV files from the KMA data portal UI."""

    name = "kma"
    page_url = "https://data.kma.go.kr/data/grnd/selectAsosRltmList.do?pgmNo=36"

    def __init__(self, config: CollectionConfig):
        self.config = config

    @staticmethod
    def _latest_available_date() -> date:
        # The portal documents that the previous day's observations are ready after 10:00.
        return date.today() - timedelta(days=1)

    def _existing_files(self) -> list[Path]:
        return sorted(self.config.existing_weather_dir.glob("OBS_ASOS_TIM_*.csv"))

    @staticmethod
    def _file_range(path: Path) -> tuple[date, date] | None:
        first = None
        last = None
        with path.open("r", encoding="cp949", errors="replace", newline="") as stream:
            reader = csv.reader(stream)
            next(reader, None)
            for row in reader:
                if len(row) >= 3 and len(row[2]) >= 10:
                    first = first or row[2][:10]
                    last = row[2][:10]
        if not first or not last:
            return None
        return date.fromisoformat(first), date.fromisoformat(last)

    def _latest_existing_date(self) -> date | None:
        latest = None
        for path in self._existing_files():
            observed = self._file_range(path)
            if observed:
                latest = max(latest or observed[1], observed[1])
        return latest

    @staticmethod
    def _chunks(start: date, end: date):
        """ASOS hourly UI permits at most one year per query."""
        current = start
        while current <= end:
            # Calendar-year chunks avoid Feb-29 replacement errors and align
            # directly with the repository's annual ASOS files.
            chunk_end = min(end, date(current.year, 12, 31))
            yield current, chunk_end
            current = chunk_end + timedelta(days=1)

    def _build_driver(self, download_dir: Path):
        from selenium import webdriver

        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", {
            "download.default_directory": str(download_dir.resolve()),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        })
        profile = os.getenv("KMA_CHROME_USER_DATA_DIR")
        if profile:
            options.add_argument(f"--user-data-dir={Path(profile).resolve()}")
        if os.getenv("KMA_HEADLESS", "0") == "1":
            options.add_argument("--headless=new")
        return webdriver.Chrome(options=options)

    @staticmethod
    def _set_search_form(driver, start: date, end: date) -> None:
        result = driver.execute_script(
            """
            const visible = el => !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
            const dateInputs = [...document.querySelectorAll('input')]
              .filter(el => visible(el) && /^\\d{8}$/.test((el.value || '').replace(/-/g,'')));
            if (dateInputs.length < 2) return {ok:false, reason:'date_inputs'};
            const setValue = (el, value) => {
              const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
              setter.call(el, value); el.dispatchEvent(new Event('input', {bubbles:true}));
              el.dispatchEvent(new Event('change', {bubbles:true}));
            };
            setValue(dateInputs[0], arguments[0]); setValue(dateInputs[1], arguments[1]);
            const selects = [...document.querySelectorAll('select')].filter(visible);
            const hourly = selects.find(s => [...s.options].some(o => /시간/.test(o.textContent)));
            if (!hourly) return {ok:false, reason:'data_type'};
            const option = [...hourly.options].find(o => /시간/.test(o.textContent));
            hourly.value = option.value; hourly.dispatchEvent(new Event('change', {bubbles:true}));
            const form = dateInputs[0].closest('form') || document;
            [...form.querySelectorAll('input[type=checkbox]')]
              .filter(el => !el.disabled).forEach(el => { if (!el.checked) el.click(); });
            return {ok:true};
            """,
            start.strftime("%Y%m%d"), end.strftime("%Y%m%d"),
        )
        if not result or not result.get("ok"):
            raise RuntimeError(f"KMA search form changed: {result}")

    @staticmethod
    def _click_text(driver, text: str) -> None:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        xpath = f"//*[self::a or self::button or self::input][contains(normalize-space(.),'{text}') or @value='{text}']"
        element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].click();", element)

    @staticmethod
    def _wait_for_download(download_dir: Path, before: set[Path], timeout: int = 180) -> Path:
        deadline = time.time() + timeout
        while time.time() < deadline:
            partials = list(download_dir.glob("*.crdownload")) + list(download_dir.glob("*.part"))
            candidates = [p for p in download_dir.iterdir() if p.is_file() and p not in before and p.suffix.lower() in {".csv", ".zip"}]
            if candidates and not partials:
                return max(candidates, key=lambda p: p.stat().st_mtime)
            time.sleep(1)
        raise TimeoutError("KMA CSV download did not complete within 180 seconds")

    def _download_chunks(self, start: date, end: date, download_dir: Path) -> list[Path]:
        download_dir.mkdir(parents=True, exist_ok=True)
        driver = self._build_driver(download_dir)
        downloaded = []
        try:
            for chunk_start, chunk_end in self._chunks(start, end):
                driver.get(self.page_url)
                if "login" in driver.current_url.lower():
                    raise RuntimeError("KMA login is required; configure KMA_CHROME_USER_DATA_DIR with a signed-in profile")
                self._set_search_form(driver, chunk_start, chunk_end)
                self._click_text(driver, "조회")
                before = set(download_dir.iterdir())
                self._click_text(driver, "CSV")
                downloaded.append(self._wait_for_download(download_dir, before))
        finally:
            driver.quit()
        return downloaded

    @staticmethod
    def _read_csv(path: Path) -> tuple[list[str], list[list[str]]]:
        # New merged files are UTF-8-sig; legacy portal files are CP949.
        for encoding in ("utf-8-sig", "utf-8", "cp949"):
            try:
                with path.open("r", encoding=encoding, newline="") as stream:
                    rows = list(csv.reader(stream))
                return (rows[0], rows[1:]) if rows else ([], [])
            except UnicodeDecodeError:
                continue
        raise ValueError(f"Unable to decode downloaded KMA CSV: {path}")

    def _merge_by_year(self, downloads: list[Path]) -> list[Path]:
        grouped: dict[int, tuple[list[str], list[list[str]]]] = {}
        for path in downloads:
            header, rows = self._read_csv(path)
            for row in rows:
                if len(row) < 3:
                    continue
                year = int(row[2][:4])
                grouped.setdefault(year, (header, []))[1].append(row)

        outputs = []
        self.config.existing_weather_dir.mkdir(parents=True, exist_ok=True)
        for year, (header, new_rows) in grouped.items():
            output = self.config.existing_weather_dir / f"OBS_ASOS_TIM_{year}.csv"
            existing_header, existing_rows = self._read_csv(output) if output.exists() else (header, [])
            final_header = existing_header or header
            unique = {tuple(row[:3]): row for row in [*existing_rows, *new_rows] if len(row) >= 3}
            ordered = sorted(unique.values(), key=lambda row: (row[0], row[2]))
            temporary = output.with_suffix(".csv.part")
            with temporary.open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.writer(stream)
                writer.writerow(final_header)
                writer.writerows(ordered)
            temporary.replace(output)
            outputs.append(output)
        return outputs

    def collect(self) -> CollectionResult:
        effective_end = min(self.config.end_date, self._latest_available_date())
        latest = self._latest_existing_date()
        start = max(self.config.start_date, latest + timedelta(days=1) if latest else self.config.start_date)
        if start > effective_end and not self.config.overwrite:
            return CollectionResult(
                self.name, "existing_data_reused", self._existing_files(),
                message=f"Existing ASOS data already covers through {latest}",
            )

        if self.config.overwrite:
            start = self.config.start_date
        temp_dir = self.config.output_dir / "kma" / "downloads"
        downloads = self._download_chunks(start, effective_end, temp_dir)
        outputs = self._merge_by_year(downloads)
        for path in downloads:
            path.unlink(missing_ok=True)
        return CollectionResult(
            self.name, "downloaded", outputs,
            message=f"KMA portal data merged through {effective_end.isoformat()}",
        )
