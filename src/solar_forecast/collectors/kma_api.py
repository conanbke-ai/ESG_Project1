"""ASOS API collection with monthly checkpoints and the existing weather/Gold contract."""
from __future__ import annotations

import calendar
from datetime import date, datetime, timedelta, timezone
import json
import os
from pathlib import Path
from uuid import uuid4

from solar_forecast.collectors.collection_config import CollectionConfig
from solar_forecast.collectors.contracts import CollectionResult
from solar_forecast.collectors.kma_api_client import ASOS_SERVICE_KEY_ENV, AsosApiError, AsosHourlyApiClient, parse_asos_page
from solar_forecast.datasets.asos_weather_store import exclusive_asos_collection, merge_asos_observations
from solar_forecast.infrastructure.artifact_store import sha256_file, write_json_atomic

ASOS_PARTITION_CONTRACT = "solar-asos-api-partition.v1"


def iter_asos_months(start: date, end: date):
    """Yield closed calendar-month intervals bounded by the requested date range."""
    while start <= end:
        last = min(end, date(start.year, start.month, calendar.monthrange(start.year, start.month)[1]))
        yield start, last
        start = last + timedelta(days=1)


def validate_asos_observations(items: list[dict], station: str, start: date, end: date) -> None:
    """Reject wrong stations, out-of-range/non-hourly times, and duplicate keys."""
    seen = set()
    valid = True
    try:
        for item in items:
            timestamp = datetime.fromisoformat(str(item["tm"]))
            if str(int(item["stnId"])) != station or not start <= timestamp.date() <= end:
                valid = False
            if timestamp.tzinfo is not None or timestamp.minute or timestamp.second or timestamp.microsecond or timestamp in seen:
                valid = False
            seen.add(timestamp)
    except (ValueError, TypeError, KeyError):
        valid = False
    if not valid:
        raise AsosApiError("ASOS returned invalid or duplicate station/hour observations")


class KmaAsosApiCollector:
    """Collect observations, then publish validated station/hour weather partitions."""
    name = "kma"

    def __init__(self, config: CollectionConfig, *, client: AsosHourlyApiClient | None = None):
        self.config = config
        self._client = client

    def collect(self) -> CollectionResult:
        key = os.getenv(ASOS_SERVICE_KEY_ENV, "").strip()
        if self._client is None and not key:
            return CollectionResult(self.name, "configuration_required", message=f"Set {ASOS_SERVICE_KEY_ENV} in .env.local using the Decoding key")
        if not self.config.station_ids:
            return CollectionResult(self.name, "configuration_required", message="ASOS API requires explicit --station-ids (for example 108,159)")
        stations = tuple(dict.fromkeys(str(int(station)) for station in self.config.station_ids))
        yesterday_kst = datetime.now(timezone(timedelta(hours=9))).date() - timedelta(days=1)
        end = min(self.config.end_date, yesterday_kst)
        if self.config.start_date > end:
            return CollectionResult(self.name, "no_data", message="ASOS provides observations only through the previous day in Korea")
        client = self._client or AsosHourlyApiClient(key, max_calls=self.config.api_max_calls)
        paths: list[Path] = []
        rows = 0
        empty = 0
        with exclusive_asos_collection(self.config.existing_weather_dir):
            for station in stations:
                for start, last in iter_asos_months(self.config.start_date, end):
                    items, artifacts = self._collect_partition(client, station, start, last)
                    paths.extend(artifacts)
                    paths.extend(merge_asos_observations(self.config.existing_weather_dir, items))
                    rows += len(items)
                    empty += not items
        return CollectionResult(self.name, "completed" if rows else "no_data", files=list(dict.fromkeys(paths)), rows=rows,
                                message=f"ASOS API: {client.calls} calls, {empty} empty partitions; end={end.isoformat()}; observations are not forecasts")

    def _collect_partition(self, client: AsosHourlyApiClient, station: str, start: date, end: date) -> tuple[list[dict], list[Path]]:
        directory = self.config.output_dir / "kma" / "asos_hourly" / f"station_{station}" / f"{start:%Y%m%d}_{end:%Y%m%d}"
        manifest_path = directory / "partition_manifest.json"
        request = {"station_id": station, "start_date": start.isoformat(), "end_date": end.isoformat()}
        if manifest_path.exists() and not self.config.overwrite:
            cached = self._read_partition(manifest_path, request, station, start, end)
            if len(cached[0]) == ((end-start).days+1)*24:
                return cached
            # A complete API response can still lack observations. Query it again
            # on the next run so late publication does not become a permanent gap.
        snapshot = directory / "responses" / uuid4().hex
        items: list[dict] = []
        pages: list[dict] = []
        total = None
        page_number = 1
        while total is None or len(items) < total:
            page = client.fetch_page(station, start, end, page_number)
            if total is not None and total != page.total_count:
                raise AsosApiError("ASOS totalCount changed during pagination; rerun this partition")
            total = page.total_count
            items.extend(page.items)
            if len(items) > total:
                raise AsosApiError("ASOS pagination exceeded totalCount")
            validate_asos_observations(items, station, start, end)
            snapshot.mkdir(parents=True, exist_ok=True)
            path = snapshot / f"page_{page_number:04d}.json"
            path.write_bytes(page.raw_bytes)
            pages.append({"path": str(path.relative_to(directory)), "sha256": sha256_file(path), "rows": len(page.items)})
            page_number += 1
        # Empty responses are retained but not marked complete: the provider may publish later.
        if items:
            write_json_atomic(manifest_path, {
                "contract": ASOS_PARTITION_CONTRACT, "schema_version": 1,
                "request": request, "rows": len(items), "pages": pages,
                "expected_hours": ((end-start).days+1)*24,
                "observed_hours": len(items),
                "complete_response": True,
                "collected_at_utc": datetime.now(timezone.utc).isoformat(),
            })
        artifacts = [directory / page["path"] for page in pages]
        if items:
            artifacts.append(manifest_path)
        return items, artifacts

    @staticmethod
    def _read_partition(manifest_path: Path, request: dict, station: str, start: date, end: date) -> tuple[list[dict], list[Path]]:
        valid = True
        items: list[dict] = []
        paths: list[Path] = []
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            if manifest["contract"] != ASOS_PARTITION_CONTRACT or manifest["schema_version"] != 1 or manifest["request"] != request or not manifest["complete_response"]:
                raise ValueError("manifest mismatch")
            directory = manifest_path.parent.resolve()
            for number, artifact in enumerate(manifest["pages"], 1):
                path = (directory / artifact["path"]).resolve()
                if not path.is_relative_to(directory) or sha256_file(path) != artifact["sha256"]:
                    raise ValueError("artifact integrity mismatch")
                page = parse_asos_page(path.read_bytes(), number)
                if page.total_count != manifest["rows"] or len(page.items) != artifact["rows"]:
                    raise ValueError("artifact count mismatch")
                items.extend(page.items)
                paths.append(path)
            if len(items) != manifest["rows"] or not items:
                raise ValueError("incomplete partition")
            validate_asos_observations(items, station, start, end)
        except (OSError, ValueError, TypeError, KeyError, AsosApiError):
            valid = False
        if not valid:
            raise AsosApiError("ASOS cache integrity check failed; inspect artifacts and rerun with --overwrite to create a fresh response snapshot")
        return items, [*paths, manifest_path]
