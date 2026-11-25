"""Calendar coverage and population tests; fixtures are not model accuracy evidence."""
from __future__ import annotations

import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

import numpy as np
import pandas as pd

from solar_forecast.config_loader import PROJECT_ROOT
from solar_forecast.evaluation.experiment_config import build_candidate_configs, configure_smoke_experiment, experiment_plan, load_experiment_config
from solar_forecast.evaluation.split_audit import audit_temporal_frame, run_split_audit
from solar_forecast.evaluation.temporal_split import TemporalBoundaries, TemporalSplitConfig, TemporalSplitter, calendar_split_for_execution


DATES = {"train_end": "2024-01-01T23:00:00", "validation_end": "2024-01-02T23:00:00",
         "calibration_end": "2024-01-03T23:00:00", "test_end": "2024-01-04T23:00:00"}


def _fixture() -> pd.DataFrame:
    hours = pd.date_range("2024-01-01", periods=100, freq="h")
    frame = pd.concat([
        pd.DataFrame({"plant_id": "known", "timestamp": hours}),
        pd.DataFrame({"plant_id": "new", "timestamp": hours[76:]}),
    ], ignore_index=True)
    frame["plant"] = frame.plant_id
    frame["region"] = "fixture_region"
    frame["energy_source"] = "solar"
    frame["generation_mwh"] = 1.0
    frame["quality_train_eligible"] = True
    frame.loc[10, "quality_train_eligible"] = False
    frame.loc[11, "generation_mwh"] = np.nan
    return frame


class CalendarSplitTests(unittest.TestCase):
    def test_calendar_is_unchanged_by_added_data_and_shared_by_plants(self):
        splitter = TemporalSplitter(TemporalSplitConfig(**DATES))
        original = pd.Series(pd.date_range("2024-01-01", periods=96, freq="h"))
        boundaries = splitter.boundaries(original)
        added = pd.concat([original, pd.Series(pd.date_range("2025-01-01", periods=24, freq="h"))], ignore_index=True)
        self.assertEqual(boundaries, splitter.boundaries(added))
        pd.testing.assert_series_equal(splitter.labels(original, boundaries), splitter.labels(added, boundaries).iloc[:96])
        self.assertTrue(splitter.labels(added, boundaries).iloc[96:].isna().all())
        plants = pd.concat([pd.DataFrame({"timestamp": original, "plant": "old"}),
                            pd.DataFrame({"timestamp": original.iloc[40:], "plant": "late"})], ignore_index=True)
        parts = splitter.split_frame(plants)
        self.assertEqual(parts.test.groupby("plant").timestamp.min().nunique(), 1)

    def test_purge_and_inclusive_test_end(self):
        splitter = TemporalSplitter(TemporalSplitConfig(**DATES, gap_hours=2))
        timestamps = pd.Series(pd.date_range("2024-01-01", periods=100, freq="h"))
        labels = splitter.labels(timestamps, splitter.boundaries(timestamps))
        self.assertEqual(labels.iloc[23], "train")
        self.assertTrue(labels.iloc[24:26].isna().all())
        self.assertEqual(labels.iloc[26], "validation")
        self.assertEqual(labels.iloc[95], "test")
        self.assertTrue(labels.iloc[96:].isna().all())

    def test_rejects_partial_unordered_timezone_and_nonfinite_config(self):
        bad = [{"train_end": DATES["train_end"]},
               {**DATES, "test_end": DATES["train_end"]},
               {**DATES, "train_end": "2024-01-01T23:00:00+09:00"},
               {**DATES, "train_end": "today"},
               {**DATES, "test_end": "not_a_date"}, {**DATES, "gap_hours": 24},
               {"validation_fraction": float("nan")}, {"gap_hours": 1.5}]
        for values in bad:
            with self.subTest(values=values), self.assertRaises(ValueError):
                TemporalSplitConfig(**values)

    def test_rejects_aware_and_mixed_timezone_data(self):
        splitter = TemporalSplitter(TemporalSplitConfig(**DATES))
        for timestamps in (pd.date_range("2024-01-01", periods=2, tz="Asia/Seoul"),
                           ["2024-01-01T00:00:00", "2024-01-01T01:00:00+09:00"]):
            with self.subTest(timestamps=timestamps), self.assertRaisesRegex(ValueError, "timezone"):
                splitter.boundaries(timestamps)

    def test_saved_boundaries_round_trip_without_recomputation(self):
        splitter = TemporalSplitter(TemporalSplitConfig(**DATES, gap_hours=2))
        bounds = splitter.boundaries(pd.Series(["2024-01-01"]))
        self.assertEqual(TemporalBoundaries.from_dict(json.loads(json.dumps(bounds.to_dict()))), bounds)
        legacy = TemporalSplitter().boundaries(pd.date_range("2024-01-01", periods=100, freq="h"))
        self.assertNotIn("test_end", legacy.to_dict())
        self.assertEqual(TemporalBoundaries.from_dict(legacy.to_dict()), legacy)

    def test_fraction_mode_remains_four_way_and_global(self):
        frame = pd.DataFrame({"timestamp": pd.date_range("2024-01-01", periods=200, freq="h")})
        split = TemporalSplitter(TemporalSplitConfig(gap_hours=4)).split_frame(frame)
        self.assertLess(split.train.timestamp.max() + pd.Timedelta(hours=4), split.validation.timestamp.min())
        self.assertEqual(split.test.timestamp.max(), frame.timestamp.max())

    def test_all_model_candidates_receive_same_frozen_dates(self):
        values = load_experiment_config(PROJECT_ROOT / "config/experiments/optimized.json")
        values["split"] = {**DATES, "purge_gap_hours": 2}
        values["models"]["cnn_bilstm"]["training_overrides"] = {"train_end": "1999-01-01", "test_fraction": 0.3}
        candidates = build_candidate_configs(values, 6, Path("fixture_run"))
        for _, candidate in candidates:
            self.assertEqual({key: candidate.values[key] for key in DATES}, DATES)
            self.assertEqual(candidate.values["purge_gap_hours"], 6)
            self.assertEqual(candidate.values["test_fraction"], 0.15)

    def test_plan_rejects_calendar_intervals_shorter_than_forecast_purge(self):
        values = load_experiment_config(PROJECT_ROOT / "config/experiments/optimized.json")
        values["split"] = {**DATES, "purge_gap_hours": 2}
        with self.assertRaisesRegex(ValueError, "purge gap"):
            build_candidate_configs(values, 72, Path("fixture_run"))

    def test_calendar_smoke_records_fraction_override_and_keeps_four_partitions(self):
        values = load_experiment_config(PROJECT_ROOT / "config/experiments/optimized.json")
        dates = {"train_end": "2023-12-31T23:00:00", "validation_end": "2024-06-30T23:00:00",
                 "calibration_end": "2024-12-31T23:00:00", "test_end": "2025-12-31T23:00:00"}
        values["split"] = {**dates, "purge_gap_hours": 168}
        standalone, reason = calendar_split_for_execution(values["split"], smoke=True)
        smoke = configure_smoke_experiment(values)
        plan = experiment_plan(smoke)
        self.assertEqual(plan["split"]["mode"], "fraction")
        self.assertEqual(plan["smoke_split_override"], reason)
        self.assertEqual(reason["requested_calendar"], dates)
        self.assertFalse(reason["accuracy_evidence"])
        self.assertEqual(values["split"]["test_end"], dates["test_end"])
        candidates = build_candidate_configs(smoke, 1, Path("fixture_smoke"))
        for _, candidate in candidates:
            fields, metadata = calendar_split_for_execution(candidate.values, smoke=True)
            self.assertEqual(fields, standalone)
            self.assertEqual(metadata, reason)
            cfg = TemporalSplitConfig.from_mapping({**candidate.values, **fields})
            splits = TemporalSplitter(cfg).split_frame(pd.DataFrame({
                "timestamp": pd.date_range("2022-01-01", periods=512, freq="h")}))
            self.assertTrue(all(len(getattr(splits, name)) > 0 for name in ("train", "validation", "calibration", "test")))


class SplitAuditTests(unittest.TestCase):
    def test_audit_accounts_for_exclusions_purge_cap_and_cold_start(self):
        frame = _fixture()
        before = frame.copy(deep=True)
        report = audit_temporal_frame(frame, TemporalSplitConfig(**DATES, gap_hours=2))
        self.assertEqual(report["population"]["eligible_observed_rows"], 122)
        self.assertEqual(report["population"]["assigned_rows"], 108)
        self.assertEqual(report["population"]["purge_gap_rows"], 6)
        self.assertEqual(report["population"]["after_test_end_rows"], 8)
        self.assertEqual(report["splits"]["test"]["rows_without_train_history"], 20)
        self.assertEqual(report["splits"]["test"]["plants_without_train_history"], ["new"])
        self.assertEqual(sum(item["rows"] for item in report["per_plant_month"]), 108)
        self.assertEqual(sum(item["rows"] for item in report["per_season"]), 108)
        self.assertFalse(report["prediction_or_training_performed"])
        pd.testing.assert_frame_equal(frame, before)

    def test_audit_exposes_empty_partitions_and_rejects_duplicate_keys(self):
        report = audit_temporal_frame(_fixture().iloc[:10], TemporalSplitConfig(**DATES))
        self.assertIn("empty_partition:test", report["warnings"])
        with self.assertRaisesRegex(ValueError, "Duplicate"):
            audit_temporal_frame(pd.concat([_fixture(), _fixture().iloc[:1]]), TemporalSplitConfig(**DATES))

    def test_data_only_command_works_and_uses_horizon_specific_purge(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            data = root / "gold.csv.gz"
            _fixture().to_csv(data, index=False)
            config = load_experiment_config(PROJECT_ROOT / "config/experiments/optimized.json")
            config.update(input_dataset=str(data), horizons_hours=[1, 6])
            config["split"] = {**DATES, "purge_gap_hours": 2}
            path = root / "experiment.json"
            path.write_text(json.dumps(config))
            output = root / "audit.json"
            command = [sys.executable, str(PROJECT_ROOT / "tools/audit_temporal_split.py"),
                       "--config", str(path), "--output", str(output)]
            completed = subprocess.run(command, capture_output=True, text=True)
            self.assertEqual(completed.returncode, 0, completed.stderr)
            report = json.loads(output.read_text())
            self.assertEqual([item["boundaries"]["gap_hours"] for item in report["audits"]], [2, 6])
            self.assertEqual(len(report["source_files"][0]["sha256"]), 64)
            with self.assertRaisesRegex(ValueError, "outside"):
                run_split_audit(path, output_path=data)


if __name__ == "__main__":
    unittest.main()
