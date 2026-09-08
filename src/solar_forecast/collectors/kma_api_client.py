"""Bounded ASOS HTTP requests, response validation, and credential-safe errors."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
import json
from typing import Callable
from urllib.parse import urlencode
from urllib.request import HTTPRedirectHandler, Request, build_opener

ASOS_HOURLY_ENDPOINT = "https://apis.data.go.kr/1360000/AsosHourlyInfoService/getWthrDataList"
ASOS_SERVICE_KEY_ENV = "KMA_ASOS_SERVICE_KEY"


class AsosApiError(RuntimeError):
    """A safe error whose message never contains credentials or response bodies."""


class _RejectRedirects(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def fetch_asos_response(url: str, timeout: float) -> bytes:
    """Read one response without forwarding the query-string key to redirects."""
    request = Request(url, headers={"Accept": "application/json"})
    with build_opener(_RejectRedirects()).open(request, timeout=timeout) as response:
        # A monthly station query is small; bound malformed server responses too.
        content = response.read(8 * 1024 * 1024 + 1)
        if len(content) > 8 * 1024 * 1024:
            raise AsosApiError("ASOS response exceeds the supported size")
        return content


@dataclass(frozen=True)
class AsosPage:
    raw_bytes: bytes
    items: tuple[dict, ...]
    total_count: int
    page_number: int


def parse_asos_page(raw_bytes: bytes, page_number: int) -> AsosPage:
    """Validate JSON pagination, including the provider's no-data result code."""
    try:
        payload = json.loads(raw_bytes)
        response = payload["response"]
        code = str(response["header"]["resultCode"])
        if code == "03":
            return AsosPage(raw_bytes, (), 0, page_number)
        if code != "00":
            # Do not interpolate resultMsg; error responses can echo request URLs.
            raise ValueError("provider rejected request")
        body = response["body"]
        total = int(body["totalCount"])
        actual_page = int(body["pageNo"])
        page_size = int(body["numOfRows"])
        container = body.get("items") or {}
        items = container.get("item", []) if isinstance(container, dict) else None
        if isinstance(items, dict):
            items = [items]
        if not isinstance(items, list) or not all(isinstance(item, dict) for item in items):
            raise ValueError("invalid items")
        if total < 0 or actual_page != page_number or page_size < 1:
            raise ValueError("invalid pagination")
        if len(items) > page_size or len(items) > total or (total and not items):
            raise ValueError("incomplete page")
        return AsosPage(raw_bytes, tuple(items), total, actual_page)
    except (ValueError, TypeError, KeyError, AttributeError, OverflowError):
        raise AsosApiError("ASOS rejected the request or returned invalid JSON/pagination; check key approval, dates, and station IDs") from None


class AsosHourlyApiClient:
    """Use one explicit call budget across all stations and monthly partitions."""

    def __init__(
        self, service_key: str, *, max_calls: int, timeout: float = 30,
        transport: Callable[[str, float], bytes] = fetch_asos_response,
    ):
        if not service_key.strip() or max_calls < 1 or timeout <= 0:
            raise ValueError("ASOS requires a key, positive call budget, and timeout")
        self._service_key = service_key.strip()
        self.max_calls = max_calls
        self.timeout = timeout
        self.calls = 0
        self._transport = transport

    def fetch_page(self, station_id: str, start: date, end: date, page_number: int) -> AsosPage:
        if self.calls >= self.max_calls:
            raise AsosApiError("ASOS API call budget exhausted; completed partitions are reusable")
        query = urlencode({
            "serviceKey": self._service_key,
            "pageNo": page_number, "numOfRows": 999, "dataType": "JSON",
            "dataCd": "ASOS", "dateCd": "HR", "stnIds": station_id,
            "startDt": start.strftime("%Y%m%d"), "startHh": "00",
            "endDt": end.strftime("%Y%m%d"), "endHh": "23",
        })
        self.calls += 1
        failed = False
        try:
            raw = self._transport(ASOS_HOURLY_ENDPOINT + "?" + query, self.timeout)
            if self._service_key.encode() in raw or query.encode() in raw:
                failed = True
        except Exception:
            # Raise outside the handler so traceback context cannot retain HTTP URLs.
            failed = True
        if failed:
            raise AsosApiError("ASOS HTTP request failed; check connectivity, key approval, and API quota")
        return parse_asos_page(raw, page_number)
