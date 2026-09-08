"""Offline ASOS boundary tests; synthetic observations never represent live coverage."""
from __future__ import annotations

import csv
from datetime import date
import json
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import traceback
import unittest
from unittest.mock import patch
from urllib.parse import parse_qs, urlsplit

from solar_forecast.cli import build_parser
from solar_forecast.collectors.collection_config import CollectionConfig
from solar_forecast.collectors.collection_service import CollectionService
from solar_forecast.collectors.collector_factory import build_collector
from solar_forecast.collectors.kma_api import KmaAsosApiCollector, iter_asos_months
from solar_forecast.collectors.kma_api_client import AsosApiError, AsosHourlyApiClient, parse_asos_page
from solar_forecast.datasets.asos_weather_store import exclusive_asos_collection, merge_asos_observations


def observation(hour=0, station="159"):
    return {"stnId": station, "stnNm": "부산", "tm": f"2025-01-01 {hour:02d}:00",
            "ta": "3.5", "rn": "", "ws": "2.1", "hm": "60", "ss": "0", "icsr": "0",
            "dc10Tca": "0", "dc10LmcsCa": "0", "taQcflg": "0"}


def response_bytes(items, *, total=None, page=1, size=999, code="00"):
    return json.dumps({"response": {"header": {"resultCode": code, "resultMsg": "NORMAL_SERVICE"},
        "body": {"totalCount": len(items) if total is None else total, "pageNo": page,
                 "numOfRows": size, "items": {"item": items}}}}, ensure_ascii=False).encode()


class AsosApiTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.config = CollectionConfig(date(2025, 1, 1), date(2025, 1, 1), station_ids=("159",),
            sources=("kma",), output_dir=self.root / "raw", existing_weather_dir=self.root / "weather", kma_mode="api")

    def client(self, responses, max_calls=10):
        self.urls = []
        iterator = iter(responses)
        def transport(url, timeout):
            self.urls.append(url)
            return next(iterator)
        return AsosHourlyApiClient("SYNTHETIC+KEY/==", max_calls=max_calls, transport=transport)

    def test_decoding_key_is_url_encoded_once_and_query_is_explicit(self):
        client = self.client([response_bytes([observation()])])
        client.fetch_page("159", self.config.start_date, self.config.end_date, 1)
        self.assertEqual(
            self.urls[0].split("?", 1)[0],
            "https://apis.data.go.kr/1360000/AsosHourlyInfoService/getWthrDataList",
        )
        params = parse_qs(urlsplit(self.urls[0]).query)
        self.assertEqual(params["serviceKey"], ["SYNTHETIC+KEY/=="])
        self.assertEqual(params["stnIds"], ["159"])
        self.assertEqual(params["startHh"], ["00"])
        self.assertEqual(params["endHh"], ["23"])
        self.assertEqual(params["dataCd"], ["ASOS"])
        self.assertEqual(params["dateCd"], ["HR"])

    def test_paginated_collection_reuses_complete_hash_checked_partition(self):
        rows = [observation(hour) for hour in range(24)]
        client = self.client([response_bytes(rows[:12], total=24, size=12), response_bytes(rows[12:], total=24, page=2, size=12)])
        collector = KmaAsosApiCollector(self.config, client=client)
        result = collector.collect()
        self.assertEqual(result.rows, 24)
        self.assertEqual(client.calls, 2)
        before = (self.root / "weather/OBS_ASOS_TIM_2025.csv").read_bytes()
        collector.collect()
        self.assertEqual(client.calls, 2)
        self.assertEqual(before, (self.root / "weather/OBS_ASOS_TIM_2025.csv").read_bytes())
        for path in (self.root / "raw").rglob("*.json"):
            self.assertNotIn(b"SYNTHETIC", path.read_bytes())

    def test_partial_observation_coverage_is_queried_again(self):
        client = self.client([response_bytes([observation()]), response_bytes([observation(), observation(1)])])
        collector = KmaAsosApiCollector(self.config, client=client)
        self.assertEqual(collector.collect().rows, 1)
        self.assertEqual(collector.collect().rows, 2)
        self.assertEqual(len(list((self.root / "raw").rglob("page_0001.json"))), 2)

    def test_budget_failure_does_not_publish_partial_weather_or_complete_manifest(self):
        client = self.client([response_bytes([observation()], total=2, size=1)], max_calls=1)
        with self.assertRaisesRegex(AsosApiError, "budget"):
            KmaAsosApiCollector(self.config, client=client).collect()
        self.assertFalse(list((self.root / "raw").rglob("partition_manifest.json")))
        self.assertFalse(list((self.root / "weather").glob("OBS_*.csv")))
        self.assertFalse((self.root / "weather/.asos_collection.lock").exists())

    def test_cache_tampering_is_rejected_without_network_or_csv_changes(self):
        client = self.client([response_bytes([observation(hour) for hour in range(24)])])
        collector = KmaAsosApiCollector(self.config, client=client)
        collector.collect()
        path = next((self.root / "raw").rglob("page_0001.json"))
        path.write_bytes(b"{}")
        with self.assertRaisesRegex(AsosApiError, "integrity"):
            collector.collect()
        self.assertEqual(client.calls, 1)

    def test_no_data_does_not_create_complete_manifest(self):
        client = self.client([response_bytes([], code="03")])
        self.assertEqual(KmaAsosApiCollector(self.config, client=client).collect().status, "no_data")
        self.assertFalse(list((self.root / "raw").rglob("partition_manifest.json")))

    def test_wrong_station_duplicate_timestamp_and_changing_total_are_rejected(self):
        cases = [
            [response_bytes([observation(station="108")])],
            [response_bytes([observation(), observation()])],
            [response_bytes([observation()], total=2, size=1), response_bytes([observation(1)], total=3, page=2)],
        ]
        for responses in cases:
            with self.subTest(responses=len(responses)):
                with self.assertRaises(AsosApiError):
                    KmaAsosApiCollector(self.config, client=self.client(responses)).collect()
        self.assertFalse(list((self.root / "weather").glob("OBS_*.csv")))

    def test_out_of_range_and_non_hourly_timestamps_are_rejected(self):
        for timestamp in ("2024-12-31 23:00", "2025-01-01 00:01", "2025-01-01T00:00:00+09:00"):
            with self.subTest(timestamp=timestamp), self.assertRaises(AsosApiError):
                KmaAsosApiCollector(self.config, client=self.client([response_bytes([{**observation(), "tm": timestamp}])])).collect()

    def test_invalid_provider_payloads_are_rejected(self):
        for raw in (b"<OpenAPI_ServiceResponse>error</OpenAPI_ServiceResponse>", b"null",
                    response_bytes([observation()], page=2), response_bytes([], total=2), response_bytes([], code="30")):
            with self.subTest(raw=raw[:20]), self.assertRaises(AsosApiError):
                parse_asos_page(raw, 1)

    def test_transport_traceback_does_not_expose_key_or_query(self):
        def transport(url, timeout):
            raise RuntimeError(url)
        client = AsosHourlyApiClient("SENSITIVE-TEST-KEY", max_calls=1, transport=transport)
        try:
            client.fetch_page("159", self.config.start_date, self.config.end_date, 1)
        except AsosApiError:
            text = traceback.format_exc()
        else:
            self.fail("transport failure was ignored")
        self.assertNotIn("SENSITIVE-TEST-KEY", text)
        self.assertNotIn("serviceKey=", text)

    def test_annual_merge_preserves_other_stations_unknown_columns_and_qc(self):
        root = self.config.existing_weather_dir
        with exclusive_asos_collection(root):
            merge_asos_observations(root, [observation(station="108"), observation()])
        path = root / "OBS_ASOS_TIM_2025.csv"
        with path.open(encoding="cp949", newline="") as stream:
            rows = list(csv.DictReader(stream))
        for row in rows:
            row["원본추가필드"] = "preserved"
        with path.open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
            writer.writeheader()
            writer.writerows(rows)
        with exclusive_asos_collection(root):
            merge_asos_observations(root, [{**observation(), "taQcflg": "1", "ta": "999"}])
        with path.open(encoding="cp949", newline="") as stream:
            merged = {row["지점"]: row for row in csv.DictReader(stream)}
        self.assertEqual(set(merged), {"108", "159"})
        self.assertEqual(merged["108"]["기온(°C)"], "3.5")
        self.assertEqual(merged["159"]["기온(°C)"], "")
        self.assertEqual(merged["159"]["taQcflg"], "1")
        self.assertEqual(merged["159"]["원본추가필드"], "preserved")
        self.assertEqual(merged["159"]["강수량(mm)"], "")

    def test_missing_api_fields_do_not_erase_existing_values(self):
        with exclusive_asos_collection(self.config.existing_weather_dir):
            merge_asos_observations(self.config.existing_weather_dir, [observation()])
            merge_asos_observations(self.config.existing_weather_dir, [{"stnId": "159", "tm": "2025-01-01 00:00", "rn": "2.3"}])
        with (self.root / "weather/OBS_ASOS_TIM_2025.csv").open(encoding="cp949", newline="") as stream:
            row = next(csv.DictReader(stream))
        self.assertEqual(row["기온(°C)"], "3.5")
        self.assertEqual(row["강수량(mm)"], "2.3")

    def test_api_and_browser_use_same_column_aware_merge(self):
        from solar_forecast.collectors.kma_browser import KmaAsosBrowserCollector
        with exclusive_asos_collection(self.config.existing_weather_dir):
            merge_asos_observations(self.config.existing_weather_dir, [observation()])
        incoming = self.root / "download.csv"
        incoming.write_text("지점,지점명,일시,기온(°C)\n108,서울,2025-01-01 00:00,1.2\n", encoding="utf-8-sig")
        with exclusive_asos_collection(self.config.existing_weather_dir):
            KmaAsosBrowserCollector(self.config)._merge_by_year([incoming])
        with (self.root / "weather/OBS_ASOS_TIM_2025.csv").open(encoding="cp949", newline="") as stream:
            rows = list(csv.DictReader(stream))
        self.assertEqual(len(rows), 2)
        self.assertIn("taQcflg", rows[0])

    def test_concurrent_weather_writer_is_rejected(self):
        with exclusive_asos_collection(self.config.existing_weather_dir):
            with self.assertRaisesRegex(RuntimeError, "lock"):
                with exclusive_asos_collection(self.config.existing_weather_dir):
                    self.fail("concurrent writer was admitted")

    def test_mode_selection_and_missing_key_have_no_implicit_fallback(self):
        from dataclasses import replace
        from solar_forecast.collectors.kma_browser import KmaAsosBrowserCollector
        with patch.dict(os.environ, {"KMA_ASOS_SERVICE_KEY": ""}):
            self.assertIsInstance(build_collector("kma", replace(self.config, kma_mode="auto")), KmaAsosBrowserCollector)
            self.assertEqual(build_collector("kma", self.config).collect().status, "configuration_required")
        with patch.dict(os.environ, {"KMA_ASOS_SERVICE_KEY": "SYNTHETIC"}):
            self.assertIsInstance(build_collector("kma", replace(self.config, kma_mode="auto")), KmaAsosApiCollector)
            self.assertIsInstance(build_collector("kma", replace(self.config, kma_mode="browser")), KmaAsosBrowserCollector)

    def test_cli_collect_and_verification_expose_the_same_station_and_mode_options(self):
        parser = build_parser()
        for command in ("collect", "verify-e2e"):
            args = parser.parse_args([command, "--start-date", "2025-01-01", "--station-ids", "108,159", "--kma-mode", "api"])
            self.assertEqual(args.station_ids, "108,159")
            self.assertEqual(args.kma_mode, "api")

    def test_collection_manifest_marks_annual_weather_as_silver(self):
        collector = KmaAsosApiCollector(self.config, client=self.client([response_bytes([observation()])]))
        service = CollectionService(self.config)
        service._factories["kma"] = lambda: collector
        result = service.run()
        self.assertEqual(result[0].rows, 1)
        manifest = json.loads(next((self.root / "raw/runs").rglob("collection_manifest.json")).read_text())
        self.assertEqual(manifest["file_artifacts"][0]["role"], "standardized_weather_silver")

    def test_month_chunks_handle_leap_year_and_cross_year(self):
        self.assertEqual(list(iter_asos_months(date(2024, 2, 28), date(2024, 3, 1))),
                         [(date(2024, 2, 28), date(2024, 2, 29)), (date(2024, 3, 1), date(2024, 3, 1))])
        self.assertEqual(len(list(iter_asos_months(date(2024, 12, 31), date(2025, 1, 1)))), 2)

    def test_asos_response_to_admission_and_gold_retains_observation_values(self):
        import pandas as pd
        import test_kospo_identity as fixtures
        from solar_forecast.collectors.generation_normalizers import DailyWideGenerationNormalizer, KOSPO_WIDE_SCHEMA
        from solar_forecast.datasets.collector_admission import CollectedGenerationAdmissionService
        from solar_forecast.datasets.plant_registry import ReviewedStationMappingCatalog
        from solar_forecast.datasets.model_dataset_builder import NationwideModelDatasetBuilder
        client = self.client([response_bytes([observation(hour) for hour in range(24)])])
        KmaAsosApiCollector(self.config, client=client).collect()
        fixtures.KospoIdentityTests.station_frame().to_csv(self.root / "weather/META_관측지점정보.csv", index=False)
        silver = self.root / "downloads"
        silver.mkdir()
        generation = DailyWideGenerationNormalizer(KOSPO_WIDE_SCHEMA).transform(fixtures.KospoIdentityTests.raw_frame(), source_file="synthetic.csv")
        generation.to_csv(silver / "synthetic_20250101.csv", index=False)
        admitted = CollectedGenerationAdmissionService(silver, self.root / "admission.json").run()
        with patch.object(ReviewedStationMappingCatalog, "from_json", return_value=fixtures.KospoIdentityTests.reviewed()):
            result = NationwideModelDatasetBuilder(self.root / "weather", fixtures.KospoIdentityTests.metadata()).build(
                self.root / "absent.csv", self.root / "gold/model_ready.csv.gz", generation_paths=admitted.accepted_paths)
        gold = pd.read_csv(result.path)
        self.assertEqual(len(gold), 24)
        self.assertEqual(set(gold["temperature_c"]), {3.5})
        self.assertEqual(gold["generation_mwh"].sum(), 1.0)


if __name__ == "__main__":
    unittest.main()
