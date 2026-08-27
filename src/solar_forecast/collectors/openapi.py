from __future__ import annotations

from datetime import date, timedelta
import os
from pathlib import Path
import xml.etree.ElementTree as ET

import pandas as pd
import requests

from solar_forecast.artifacts.manifest import write_json_atomic

from .base import CollectionResult
from .config import CollectionConfig


KOMIPO_DISCOVERY_COLUMNS = [
    "query_date",
    "station_code",
    "site_name",
    "unit_name",
    "measured_at",
    "generation_value",
    "source_unit",
]


class KomipoRenewableCollector:
    """Incrementally stage KOMIPO renewable measurements for admission review.

    The official API does not declare the unit of ``daypower``.  Consequently
    this collector writes a Bronze discovery contract and never labels values
    as MWh or feeds them directly into training.  Each station/day partition is
    atomic, resumable, and small enough to keep memory bounded.
    """

    name = "komipo"
    endpoint = "https://apis.data.go.kr/B552521/renewEnergy/getData"
    service_key_environment = "DATA_GO_SERVICE_KEY"
    page_size = 100

    def __init__(self, config: CollectionConfig, session: requests.Session | None = None):
        self.config = config
        self.session = session or requests.Session()

    def collect(self) -> CollectionResult:
        key = os.getenv(self.service_key_environment, "").strip()
        codes = tuple(dict.fromkeys(str(code).strip() for code in self.config.komipo_station_codes if str(code).strip()))
        if not key:
            return CollectionResult(
                self.name,
                "configuration_required",
                message=f"Set {self.service_key_environment} after approving API 15084511",
            )
        if not codes:
            return CollectionResult(
                self.name,
                "configuration_required",
                message="Provide --komipo-station-codes; the official endpoint requires a headquarters code",
            )

        days = (self.config.end_date - self.config.start_date).days + 1
        minimum_calls = days * len(codes)
        if minimum_calls > self.config.api_max_calls:
            return CollectionResult(
                self.name,
                "configuration_required",
                message=(
                    f"Requested range needs at least {minimum_calls} calls, above the configured "
                    f"budget {self.config.api_max_calls}; narrow the date range or raise --api-max-calls"
                ),
            )

        files: list[Path] = []
        rows = 0
        calls = 0
        for station_code in codes:
            for current in self._dates(self.config.start_date, self.config.end_date):
                destination = self._destination(station_code, current)
                empty_marker = destination.with_suffix(destination.suffix + ".empty.json")
                if not self.config.overwrite and (destination.exists() or empty_marker.exists()):
                    files.append(destination if destination.exists() else empty_marker)
                    continue
                records, used_calls = self._fetch_day(key, station_code, current)
                calls += used_calls
                if calls > self.config.api_max_calls:
                    raise RuntimeError("KOMIPO pagination exceeded the configured API call budget")
                destination.parent.mkdir(parents=True, exist_ok=True)
                if not records:
                    write_json_atomic(
                        empty_marker,
                        {
                            "query_date": current.isoformat(),
                            "station_code": station_code,
                            "status": "official_api_returned_no_rows",
                        },
                    )
                    files.append(empty_marker)
                    continue
                frame = pd.DataFrame(records, columns=KOMIPO_DISCOVERY_COLUMNS)
                frame = frame.drop_duplicates(
                    ["station_code", "site_name", "unit_name", "measured_at"],
                    keep="last",
                ).sort_values(["measured_at", "site_name", "unit_name"], kind="stable")
                temporary = destination.with_name(destination.name + ".tmp")
                frame.to_csv(
                    temporary,
                    index=False,
                    encoding="utf-8-sig",
                    compression={"method": "gzip", "compresslevel": 1, "mtime": 1},
                )
                temporary.replace(destination)
                files.append(destination)
                rows += len(frame)
        return CollectionResult(
            self.name,
            "downloaded",
            files,
            rows=rows,
            message=(
                f"Staged {len(files)} bounded station/day partitions using {calls} API calls; "
                "source generation unit remains quarantined for review"
            ),
        )

    def _fetch_day(
        self,
        service_key: str,
        station_code: str,
        current: date,
    ) -> tuple[list[dict[str, object]], int]:
        page = 1
        total = 1
        records: list[dict[str, object]] = []
        calls = 0
        while (page - 1) * self.page_size < total:
            response = self.session.get(
                self.endpoint,
                params={
                    "ServiceKey": service_key,
                    "pageNo": page,
                    "numOfRows": self.page_size,
                    "stationName": station_code,
                    "dataDate": current.strftime("%Y%m%d"),
                    "dataTerm": "DAILY",
                },
                timeout=60,
            )
            calls += 1
            response.raise_for_status()
            root = ET.fromstring(response.content)
            result_code = self._text(root, ".//resultCode")
            if result_code not in {None, "00"}:
                message = self._text(root, ".//resultMsg") or "unknown error"
                raise RuntimeError(f"KOMIPO API error {result_code}: {message}")
            total = int(self._text(root, ".//totalCount") or 0)
            for item in root.findall(".//item"):
                records.append(
                    {
                        "query_date": current.isoformat(),
                        "station_code": station_code,
                        "site_name": self._text(item, "siteterm"),
                        "unit_name": self._text(item, "unitterm"),
                        "measured_at": self._text(item, "gathdtm"),
                        "generation_value": pd.to_numeric(
                            self._text(item, "daypower"), errors="coerce"
                        ),
                        "source_unit": "portal_unspecified",
                    }
                )
            page += 1
        return records, calls

    def _destination(self, station_code: str, current: date) -> Path:
        return (
            self.config.output_dir
            / self.name
            / f"station={station_code}"
            / f"year={current.year}"
            / f"date={current.strftime('%Y%m%d')}.csv.gz"
        )

    @staticmethod
    def _dates(start: date, end: date):
        current = start
        while current <= end:
            yield current
            current += timedelta(days=1)

    @staticmethod
    def _text(root: ET.Element, path: str) -> str | None:
        element = root.find(path)
        return element.text.strip() if element is not None and element.text else None
