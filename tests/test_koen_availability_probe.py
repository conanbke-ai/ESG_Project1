from datetime import date
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path
from unittest.mock import Mock

import pytest
from selenium.common.exceptions import UnexpectedAlertPresentException, WebDriverException


SPEC = spec_from_file_location("probe_koen", Path(__file__).resolve().parents[1] / "tools/probe_koen_availability.py")
probe = module_from_spec(SPEC)
SPEC.loader.exec_module(probe)
WINDOW = (date(2026, 11, 1), date(2026, 11, 25))


@pytest.mark.parametrize(
    "dates, empty, refreshed, status",
    [(["2026-11-03 00:00:00"], [], True, "visible_in_window_dates"),
     (["2026.10.31"], [], True, "stale_or_out_of_window_results"),
     ([], ["조회된 데이터가 없습니다"], True, "provider_displayed_no_rows"),
     ([], ["조회된 데이터가 없습니다"], False, "result_refresh_unconfirmed"),
     ([], [], True, "unrecognized_result")],
)
def test_query_outcome_needs_fresh_matching_visible_evidence(dates, empty, refreshed, status):
    snapshot = {"dates": dates, "emptyMessages": empty, "fields": ["20261101", "2026-11-25"]}
    assert probe._result(snapshot, WINDOW, refreshed=refreshed)["status"] == status


def test_navigation_error_and_form_alert_do_not_mean_empty_month():
    driver = Mock()
    browser = Mock()
    browser._open.side_effect = WebDriverException("net::ERR_TUNNEL_CONNECTION_FAILED")
    unavailable = probe._probe(driver, browser, *WINDOW, 1)
    assert unavailable["status"] == "navigation_failed"
    assert unavailable["detail"] == "net::ERR_TUNNEL_CONNECTION_FAILED"
    browser.submit_query.assert_not_called()

    browser._open.side_effect = None
    browser.submit_query.side_effect = UnexpectedAlertPresentException(alert_text="조회 종료일을 확인해 주세요")
    rejected = probe._probe(driver, browser, *WINDOW, 1)
    assert rejected["status"] == "provider_alert"
    assert rejected["stage"] == "form_submission"
    assert rejected["alert_text"] == "조회 종료일을 확인해 주세요"
    assert [call.args for call in driver.execute_script.call_args_list] == [(probe.RESULT_SNAPSHOT,)]
