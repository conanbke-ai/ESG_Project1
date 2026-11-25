"""Query the public KOEN date form and record visible results without downloading CSVs."""
from __future__ import annotations

import argparse
from datetime import date, datetime, timedelta, timezone
import json
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from solar_forecast.collectors.koen_browser import KoenBrowserDownloader
from solar_forecast.infrastructure.artifact_store import write_json_atomic


DEFAULT_WINDOWS = ("2026-10-31:2026-10-31", "2026-11-01:2026-11-25", "2026-11-01:2026-11-30")
RESULT_SNAPSHOT = r"""
const dates = [], emptyMessages = [];
const dateHeader = /^(일자|일시|날짜|발전일자|발전일시|발전일|기준일자)$/;
const dateText = /^\d{4}[-./]?\d{2}[-./]?\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?$/;
const emptyMessage = /(?:조회\s*(?:된|결과)|등록\s*된|해당).*?(?:데이터|자료|내역|결과|정보).*?없/;
for (const table of document.querySelectorAll('table')) {
  if (!table.getClientRects().length) continue;
  const columns = new Set(Array.from(table.querySelectorAll('th'))
    .filter(th => dateHeader.test(th.innerText.replace(/\s+/g, '')))
    .map(th => th.cellIndex));
  const generationTable = columns.size || /발전량|발전구분|발전소/.test(
    Array.from(table.querySelectorAll('caption, th')).map(cell => cell.innerText).join(' '));
  if (!generationTable) continue;
  for (const row of table.rows) {
    if (!row.getClientRects().length || !row.querySelector('td')) continue;
    const message = row.innerText.trim();
    if (emptyMessage.test(message)) emptyMessages.push(message);
    for (const cell of row.cells) {
      const value = cell.innerText.trim();
      if (cell.tagName === 'TD' && (columns.has(cell.cellIndex) || !columns.size) && dateText.test(value)) dates.push(value);
    }
  }
}
return {dates, emptyMessages, fields: ['strDateS', 'strDateE']
  .map(id => document.getElementById(id)?.value ?? '')};
"""


def _window(value: str) -> tuple[date, date]:
    try:
        first, last = map(date.fromisoformat, value.split(":"))
        if first <= last:
            return first, last
    except ValueError:
        pass
    raise argparse.ArgumentTypeError("Use an inclusive ISO START:END date window")


def _days(texts: list[str]) -> list[date]:
    days = set()
    for text in texts:
        if match := re.fullmatch(r"(\d{4})[-./]?(\d{2})[-./]?(\d{2})(?:[ T]\d{2}:\d{2}(?::\d{2})?)?", text):
            try:
                days.add(date(*map(int, match.groups())))
            except ValueError:
                continue
    return sorted(days)


def _result(snapshot: dict, expected: tuple[date, date], *, refreshed: bool) -> dict:
    first, last = expected
    days = _days(snapshot["dates"])
    filters_match = [re.sub(r"\D", "", value) for value in snapshot["fields"]] == [f"{first:%Y%m%d}", f"{last:%Y%m%d}"]
    if days and any(not first <= day <= last for day in days):
        status = "stale_or_out_of_window_results"
    elif not refreshed:
        status = "result_refresh_unconfirmed"
    elif not filters_match:
        status = "result_filter_mismatch"
    elif days and not snapshot["emptyMessages"]:
        status = "visible_in_window_dates"
    elif snapshot["emptyMessages"] and not days:
        status = "provider_displayed_no_rows"
    else:
        status = "unrecognized_result"
    return {"status": status, "visible_dates": [day.isoformat() for day in days],
            "refresh_observed": refreshed, "displayed_filters": snapshot["fields"],
            "empty_message": " ".join(snapshot["emptyMessages"])[:250] or None}


def _probe(driver, browser, first: date, last: date, timeout: int) -> dict:
    from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException, WebDriverException
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.expected_conditions import staleness_of
    from selenium.webdriver.support.ui import WebDriverWait

    outcome = {"start_date": first.isoformat(), "end_date": last.isoformat()}
    stage = "navigation"
    try:
        browser._open(driver)
        old_document = driver.find_element(By.TAG_NAME, "html")
        before = driver.execute_script(RESULT_SNAPSHOT)
        stage = "form_submission"
        browser.submit_query(driver, first, last)
        stage = "result_read"
        snapshot, refreshed = before, False

        def read_result(current_driver):
            nonlocal snapshot, refreshed
            snapshot = current_driver.execute_script(RESULT_SNAPSHOT)
            refreshed = staleness_of(old_document)(current_driver) or any(
                snapshot[key] != before[key] for key in ("dates", "emptyMessages")
            )
            return refreshed and (bool(_days(snapshot["dates"])) or bool(snapshot["emptyMessages"]))

        try:
            WebDriverWait(driver, timeout).until(read_result)
        except TimeoutException:
            pass
        return {**outcome, **_result(snapshot, (first, last), refreshed=refreshed)}
    except UnexpectedAlertPresentException as error:
        return {**outcome, "status": "provider_alert", "stage": stage,
                "alert_text": " ".join((error.alert_text or "").split())[:250]}
    except WebDriverException as error:
        code = re.search(r"net::ERR_[A-Z_]+", error.msg or "")
        return {**outcome, "status": f"{stage}_failed", "detail": code[0] if code else type(error).__name__}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True, help="New JSON report path")
    parser.add_argument("--window", action="append", type=_window, help="Repeatable inclusive ISO START:END date range")
    parser.add_argument("--timeout", type=int, default=15, help="Page and result deadline in seconds")
    args = parser.parse_args()
    if args.output.exists() or args.timeout < 1:
        parser.error("Use a new output path and a positive timeout")
    windows = args.window or list(map(_window, DEFAULT_WINDOWS))
    observed_at = datetime.now(timezone(timedelta(hours=9))).isoformat()
    with TemporaryDirectory(prefix="koen-query-") as temporary:
        browser = KoenBrowserDownloader(Path(temporary))
        try:
            driver = browser._build_driver(Path(temporary))
        except Exception as error:
            results = [{"start_date": first.isoformat(), "end_date": last.isoformat(),
                        "status": "browser_start_failed", "detail": type(error).__name__} for first, last in windows]
        else:
            try:
                driver.set_page_load_timeout(args.timeout)
                results = [_probe(driver, browser, first, last, args.timeout) for first, last in windows]
            finally:
                driver.quit()
    report = {"contract": "solar-koen-availability-probe.v1", "observed_at_kst": observed_at,
              "source_url": KoenBrowserDownloader.page_url, "csv_download_requested": False,
              "coverage_verified": False, "windows": results}
    write_json_atomic(args.output, report)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
