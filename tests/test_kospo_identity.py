"""Synthetic regression cases for collector identity; not production row counts."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import pandas as pd

from solar_forecast.datasets.collector_admission import CollectedGenerationAdmissionService
from solar_forecast.collectors.plant_identity import KOSPO_IDENTITY_RULE
from solar_forecast.collectors.plant_identity import read_generation_partition
from solar_forecast.collectors.plant_identity import resolve_generation_identity
from solar_forecast.collectors.plant_metadata import PlantMetadata
from solar_forecast.collectors.plant_metadata import PlantMetadataCatalog
from solar_forecast.collectors.generation_normalizers import DailyWideGenerationNormalizer
from solar_forecast.collectors.generation_normalizers import GENERATION_COLUMNS
from solar_forecast.collectors.generation_normalizers import KOSPO_WIDE_SCHEMA
from solar_forecast.datasets.plant_registry import KmaStationCatalog
from solar_forecast.datasets.plant_registry import NationwidePlantRegistryBuilder
from solar_forecast.datasets.plant_registry import ReviewedStationMapping
from solar_forecast.datasets.plant_registry import ReviewedStationMappingCatalog
from solar_forecast.datasets.model_dataset_builder import NationwideModelDatasetBuilder
from solar_forecast.features.asos_features import WEATHER_COLUMN_MAP


class KospoIdentityTests(unittest.TestCase):
    def setUp(self):
        self.temporary = TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)

    @staticmethod
    def raw_frame(generator="부산 철도태양광 #2", meter="KPX"):
        row = {
            "거래일자": "2025-01-01",
            "발전소명": "신재생사업본부",
            "발전기명": generator,
            "계량구분": meter,
            **{f"{hour}시": 0.0 for hour in range(1, 25)},
        }
        row["13시"] = 1000.0
        return pd.DataFrame([row])

    def silver(self, *, legacy=False, generator="부산 철도태양광 #2"):
        frame = DailyWideGenerationNormalizer(KOSPO_WIDE_SCHEMA).transform(
            self.raw_frame(generator), source_file="synthetic.csv"
        )
        if legacy and generator == "부산 철도태양광 #2":
            frame["plant"] = "신재생사업본부"
            frame["plant_id"] = "kospo:신재생사업본부#부산 철도태양광 #2#KPX"
            frame.attrs.clear()
        return frame

    def write_silver(self, frame, name="sample_20250907.csv"):
        path = self.root / "file" / "standardized" / "downloads" / "kospo" / name
        path.parent.mkdir(parents=True, exist_ok=True)
        frame.to_csv(path, index=False, encoding="utf-8-sig")
        return path

    @staticmethod
    def metadata():
        # The public specification table leaves both capacities blank.
        return PlantMetadataCatalog([
            PlantMetadata("kospo", "부산철도 1", "solar", None, None,
                          "부산광역시 부산진구 신천대로 215"),
            PlantMetadata("kospo", "부산철도 2", "solar", None, None,
                          "부산광역시 부산진구 신천대로 215"),
        ])

    @staticmethod
    def station_frame():
        return pd.DataFrame([{
            "지점": 159, "시작일": "1904-04-09", "종료일": None,
            "지점명": "부산", "지점주소": "부산광역시 중구 대청동1가",
            "위도": 35.1047, "경도": 129.032, "노장해발고도(m)": 69.56,
        }])

    @staticmethod
    def reviewed():
        # A test-only review tests the positive gate. Production config is unchanged.
        return ReviewedStationMappingCatalog([ReviewedStationMapping(
            "kospo", "부산철도 2", 159,
            "https://example.invalid/test-only-weather-review",
            "Synthetic approval for testing the pipeline boundary only.",
        )])

    def registry(self, paths, *, reviewed=False):
        station_path = self.root / "META_관측지점정보.csv"
        self.station_frame().to_csv(station_path, index=False)
        return NationwidePlantRegistryBuilder(
            self.metadata(), KmaStationCatalog.from_metadata(station_path),
            self.reviewed() if reviewed else ReviewedStationMappingCatalog(),
        ).build(paths, self.root / "registry.csv")

    def test_new_collector_uses_generator_identity_without_changing_values(self):
        raw = self.raw_frame()
        original = raw.copy(deep=True)
        frame = DailyWideGenerationNormalizer(KOSPO_WIDE_SCHEMA).transform(raw)
        self.assertEqual(list(frame.columns), GENERATION_COLUMNS)
        self.assertEqual(set(frame["plant"]), {"부산철도 2"})
        self.assertEqual(set(frame["plant_id"]), {"kospo:부산철도 2#부산 철도태양광 #2#KPX"})
        self.assertEqual(set(frame["unit"]), {"부산 철도태양광 #2"})
        self.assertEqual(len(frame), 24)
        self.assertEqual(frame["generation_mwh"].sum(), 1.0)
        self.assertEqual(frame.loc[frame["generation_mwh"].gt(0), "timestamp"].iloc[0].hour, 12)
        self.assertTrue(frame["capacity_mw"].isna().all())
        pd.testing.assert_frame_equal(raw, original)

    def test_other_generators_are_not_mapped_by_department(self):
        for generator in ("부산 철도태양광 #1", "부산 철도태양광 #4", "다른 발전소 #2"):
            with self.subTest(generator=generator):
                frame = self.silver(generator=generator)
                self.assertEqual(set(frame["plant"]), {"신재생사업본부"})

    def test_generator_whitespace_is_normalized_but_suffixes_are_not_guessed(self):
        frame = self.silver(generator="부산철도태양광#2")
        self.assertEqual(set(frame["plant"]), {"부산철도 2"})
        for label in ("부산철도태양광#20", "부산철도태양광#2예비"):
            self.assertEqual(set(self.silver(generator=label)["plant"]), {"신재생사업본부"})

    def test_company_and_department_are_both_required(self):
        original = self.silver(legacy=True)
        for column, value in (("company", "koen"), ("plant", "다른본부"),
                              ("unit", None), ("energy_source", "wind")):
            with self.subTest(column=column):
                frame = original.copy()
                frame[column] = value
                pd.testing.assert_frame_equal(resolve_generation_identity(frame), frame)

    def test_repeated_resolution_is_idempotent_and_preserves_meters(self):
        raw = pd.concat([self.raw_frame(meter="KPX"), self.raw_frame(meter="TEST")])
        frame = DailyWideGenerationNormalizer(KOSPO_WIDE_SCHEMA).transform(raw)
        pd.testing.assert_frame_equal(resolve_generation_identity(frame), frame)
        self.assertEqual(frame["plant_id"].nunique(), 2)
        self.assertEqual(frame["generation_mwh"].sum(), 2.0)

    def test_legacy_admission_records_rule_and_preserves_file_hash(self):
        path = self.write_silver(self.silver(legacy=True))
        original = path.read_bytes()
        result = CollectedGenerationAdmissionService(path.parent, self.root / "admission.json").run()
        self.assertEqual(result.accepted_paths, (path,))
        self.assertEqual(result.accepted_rows, 24)
        self.assertEqual(path.read_bytes(), original)
        payload = json.loads(result.manifest_path.read_text())
        entry = payload["files"][0]
        self.assertEqual(entry["source_sha256"], hashlib.sha256(original).hexdigest())
        self.assertEqual(entry["identity_resolutions"][0]["rule"], KOSPO_IDENTITY_RULE)
        self.assertEqual(entry["identity_resolutions"][0]["rows"], 24)

    def test_conflicting_id_is_rejected_by_admission(self):
        frame = self.silver(legacy=True)
        frame["plant_id"] = "kospo:unrelated#2"
        path = self.write_silver(frame)
        result = CollectedGenerationAdmissionService(path.parent, self.root / "admission.json").run()
        self.assertEqual(result.rejected_count, 1)
        self.assertIn("conflicts", result.files[0].reason)

    def test_identity_aliases_cannot_double_count_one_meter_in_one_file(self):
        variants = pd.concat([
            self.raw_frame(), self.raw_frame(generator="부산철도태양광#2")
        ], ignore_index=True)
        with self.assertRaisesRegex(ValueError, "duplicate hourly keys"):
            DailyWideGenerationNormalizer(KOSPO_WIDE_SCHEMA).transform(variants)

    def test_stale_generator_id_cannot_be_silently_reassigned(self):
        frame = self.silver(legacy=True)
        frame["plant_id"] = "kospo:신재생사업본부#부산 철도태양광 #1#KPX"
        with self.assertRaisesRegex(ValueError, "conflicts"):
            resolve_generation_identity(frame)

    def test_historical_partitions_without_unit_are_unchanged(self):
        frame = self.silver().drop(columns="unit")
        path = self.write_silver(frame)
        actual = read_generation_partition(path, ("company", "plant", "generation_mwh"))
        expected = pd.read_csv(path)[list(actual.columns)]
        pd.testing.assert_frame_equal(actual, expected)

    def test_read_projection_still_rejects_missing_required_columns(self):
        path = self.write_silver(self.silver().drop(columns="generation_mwh"))
        with self.assertRaisesRegex(ValueError, "generation_mwh"):
            read_generation_partition(path, ("plant", "generation_mwh"))

    def test_real_address_does_not_bypass_weather_gate_or_infer_capacity(self):
        path = self.write_silver(self.silver(legacy=True))
        registry = self.registry([path])
        self.assertEqual(registry.iloc[0]["plant"], "부산철도 2")
        self.assertEqual(registry.iloc[0]["address"], "부산광역시 부산진구 신천대로 215")
        self.assertTrue(pd.isna(registry.iloc[0]["capacity_mw"]))
        self.assertEqual(registry.iloc[0]["model_ready_status"], "quarantined")
        with self.assertRaisesRegex(ValueError, "auditable KMA"):
            NationwideModelDatasetBuilder(self.root, self.metadata())._from_standardized_generation([path], registry)

    def test_approved_legacy_and_new_snapshots_reconcile_without_double_count(self):
        old = self.write_silver(self.silver(legacy=True), "old_20250907.csv")
        new_frame = self.silver()
        new_frame["generation_mwh"] *= 1.5
        new = self.write_silver(new_frame, "new_20250908.csv")
        paths = [old, new]
        registry = self.registry(paths, reviewed=True)
        self.assertEqual(len(registry), 1)
        builder = NationwideModelDatasetBuilder(self.root, self.metadata())
        generation = builder._from_standardized_generation(paths, registry)
        self.assertEqual(len(generation), 24)
        self.assertEqual(generation["generation_mwh"].sum(), 1.5)
        self.assertEqual(set(generation["plant_id"]), {"kospo:부산철도 2"})
        self.assertEqual([x["retained_rows_after_revision_selection"] for x in builder._source_contribution], [0, 24])

    def test_synthetic_admission_to_gold_build_with_explicit_weather_review(self):
        path = self.write_silver(self.silver(legacy=True))
        before = path.read_bytes()
        admission = CollectedGenerationAdmissionService(path.parent, self.root / "admission.json").run()
        weather_root = self.root / "weather"
        weather_root.mkdir()
        self.station_frame().to_csv(weather_root / "META_관측지점정보.csv", index=False)
        weather = pd.DataFrame({
            key: [0.0] * 24 for key in WEATHER_COLUMN_MAP
        })
        weather["지점"] = 159
        weather["지점명"] = "부산"
        weather["일시"] = pd.date_range("2025-01-01", periods=24, freq="h")
        weather.to_csv(weather_root / "OBS_ASOS_TIM_2025.csv", index=False)
        with patch.object(ReviewedStationMappingCatalog, "from_json", return_value=self.reviewed()):
            result = NationwideModelDatasetBuilder(weather_root, self.metadata()).build(
                self.root / "absent_legacy.csv", self.root / "gold" / "model_ready.csv.gz",
                generation_paths=admission.accepted_paths,
            )
        gold = pd.read_csv(result.path)
        manifest = json.loads(result.manifest_path.read_text())
        self.assertEqual(len(gold), 24)
        self.assertEqual(gold["generation_mwh"].sum(), 1.0)
        self.assertEqual(set(gold["plant_id"]), {"kospo:부산철도 2"})
        self.assertEqual(manifest["gold_source_contribution"][0]["retained_rows_after_revision_selection"], 24)
        self.assertEqual(manifest["contract"], "solar-model-ready-manifest.v1")
        self.assertTrue(list(result.partitions_dir.rglob("*.csv.gz")))
        self.assertEqual(path.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
