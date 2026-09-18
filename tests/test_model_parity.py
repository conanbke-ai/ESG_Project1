"""Strict artifact comparisons must never manufacture model performance evidence."""
from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

import pandas as pd

from solar_forecast.evaluation.model_parity import compare_prediction_files


class PredictionParityTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.baseline = self.root / "baseline.csv"
        self.candidate = self.root / "candidate.csv"
        self.frame = pd.DataFrame({
            "plant_id": ["001", "001", "002", "002"],
            "timestamp": ["2025-01-01 00:00", "2025-01-01 01:00"] * 2,
            "region": ["부산", "부산", "서울", "서울"],
            "split": ["test"] * 4,
            "y_true": [0.0, 2.0, 1.0, 3.0],
            "y_pred": [0.0, 2.0, 1.0, 3.0],
        })
        self.frame.to_csv(self.baseline, index=False, encoding="utf-8-sig")
        self.write_candidate(self.frame)

    def write_candidate(self, frame):
        frame.to_csv(self.candidate, index=False, encoding="utf-8-sig")

    def compare(self, **kwargs):
        return compare_prediction_files(self.baseline, self.candidate, **kwargs)

    def test_reordered_identical_rows_pass_and_preserve_string_ids(self):
        self.write_candidate(self.frame.iloc[::-1])
        report = self.compare(atol=0, rtol=0)
        self.assertEqual(report["status"], "passed")
        self.assertEqual(report["overall"]["baseline"]["r2"], 1)
        self.assertEqual([item["plant_id"] for item in report["per_plant"]], ["001", "002"])
        self.assertEqual(report["inputs"]["baseline"]["rows"], 4)
        self.assertEqual(report["inputs"]["baseline"]["sha256"], sha256(self.baseline.read_bytes()).hexdigest())
        self.assertEqual(report["inputs"]["candidate"]["sha256"], sha256(self.candidate.read_bytes()).hexdigest())
        self.assertFalse(report["production_accuracy_verified"])
        self.assertTrue(all(item["status"] == "unavailable" for item in report["contextual_metrics"].values()))
        json.dumps(report, allow_nan=False)

    def test_changed_predictions_fail_with_pooled_metrics_and_group_deltas(self):
        changed = self.frame.copy()
        changed["y_pred"] += 1
        self.write_candidate(changed)
        report = self.compare()
        self.assertEqual(report["status"], "failed")
        self.assertEqual(report["overall"]["prediction_difference"], {
            "max_absolute": 1.0, "mean_absolute": 1.0, "rows_outside_tolerance": 4,
        })
        self.assertEqual(report["overall"]["candidate"]["mae"], 1)
        self.assertEqual(report["overall"]["candidate"]["rmse"], 1)
        self.assertAlmostEqual(report["overall"]["candidate"]["r2"], 0.2)
        self.assertAlmostEqual(report["overall"]["delta"]["r2"], -0.8)
        self.assertEqual(len(report["per_region"]["groups"]), 2)
        self.assertTrue(all(item["delta"]["mae"] == 1 for item in report["per_plant"]))

    def test_tolerance_is_applied_against_baseline(self):
        changed = self.frame.copy()
        changed.loc[0, "y_pred"] = 5e-7
        self.write_candidate(changed)
        self.assertEqual(self.compare()["status"], "passed")
        self.assertEqual(self.compare(atol=0, rtol=0)["status"], "failed")

    def test_relative_tolerance_cannot_use_candidate_as_reference(self):
        self.frame["y_pred"] = 1
        self.frame.to_csv(self.baseline, index=False)
        changed = self.frame.copy()
        changed["y_pred"] = 2
        self.write_candidate(changed)
        self.assertEqual(self.compare(atol=0, rtol=0.75)["status"], "failed")

    def test_invalid_tolerances_are_rejected(self):
        for argument in ("atol", "rtol"):
            for value in (-1, float("inf"), float("nan"), True, "0.1"):
                with self.subTest(argument=argument, value=value):
                    with self.assertRaisesRegex(ValueError, "finite nonnegative"):
                        self.compare(**{argument: value})

    def test_duplicate_keys_are_rejected(self):
        self.write_candidate(pd.concat([self.frame, self.frame.iloc[:1]], ignore_index=True))
        with self.assertRaisesRegex(ValueError, "duplicate plant_id/timestamp"):
            self.compare()

    def test_normalized_duplicate_timestamps_are_rejected(self):
        changed = self.frame.copy()
        changed.loc[1, "timestamp"] = "2025-01-01T00:00:00"
        self.write_candidate(changed)
        with self.assertRaisesRegex(ValueError, "duplicate plant_id/timestamp"):
            self.compare()

    def test_missing_rows_and_same_count_different_keys_are_rejected(self):
        self.write_candidate(self.frame.iloc[:-1])
        with self.assertRaisesRegex(ValueError, "key sets differ"):
            self.compare()
        changed = self.frame.copy()
        changed.loc[0, "plant_id"] = "003"
        self.write_candidate(changed)
        with self.assertRaisesRegex(ValueError, "missing candidate keys=1, extra candidate keys=1"):
            self.compare()

    def test_different_observations_are_rejected_even_with_loose_tolerance(self):
        changed = self.frame.copy()
        changed.loc[0, "y_true"] += 1e-8
        self.write_candidate(changed)
        with self.assertRaisesRegex(ValueError, "observed y_true"):
            self.compare(atol=1, rtol=1)

    def test_nonfinite_or_nonnumeric_targets_and_predictions_are_rejected(self):
        for column in ("y_true", "y_pred"):
            for value in ("nan", "inf", "-inf", "", "wrong"):
                with self.subTest(column=column, value=value):
                    changed = self.frame.astype(object)
                    changed.loc[0, column] = value
                    self.write_candidate(changed)
                    with self.assertRaisesRegex(ValueError, "finite numeric"):
                        self.compare()

    def test_region_disagreement_or_missing_column_is_rejected(self):
        changed = self.frame.copy()
        changed.loc[0, "region"] = "제주"
        self.write_candidate(changed)
        with self.assertRaisesRegex(ValueError, "region labels disagree"):
            self.compare()
        self.write_candidate(self.frame.drop(columns="region"))
        with self.assertRaisesRegex(ValueError, "region column presence"):
            self.compare()

    def test_absent_regions_are_explicitly_unavailable(self):
        frame = self.frame.drop(columns="region")
        frame.to_csv(self.baseline, index=False)
        self.write_candidate(frame)
        self.assertEqual(self.compare()["per_region"]["status"], "unavailable")

    def test_split_mismatch_is_rejected(self):
        changed = self.frame.copy()
        changed["split"] = "validation"
        self.write_candidate(changed)
        with self.assertRaisesRegex(ValueError, "split labels disagree"):
            self.compare()

    def test_malformed_and_non_hourly_timestamps_are_rejected(self):
        for value in ("0", "2025-01-01", "2025-13-01 00:00", "2025-01-01 00:30", "2025-01-01 00:00:00.000000001"):
            with self.subTest(value=value):
                changed = self.frame.copy()
                changed.loc[0, "timestamp"] = value
                self.write_candidate(changed)
                with self.assertRaisesRegex(ValueError, "timestamp"):
                    self.compare()

    def test_equivalent_aware_timestamps_align_without_guessing_naive_timezone(self):
        baseline = self.frame.copy()
        baseline["timestamp"] += "+09:00"
        baseline.to_csv(self.baseline, index=False)
        changed = self.frame.copy()
        changed["timestamp"] = ["2024-12-31T15:00Z", "2024-12-31T16:00Z"] * 2
        self.write_candidate(changed)
        self.assertEqual(self.compare()["status"], "passed")
        self.write_candidate(self.frame)
        with self.assertRaisesRegex(ValueError, "timezone modes differ"):
            self.compare()

    def test_mixed_timezone_modes_are_rejected(self):
        changed = self.frame.copy()
        changed.loc[0, "timestamp"] += "+09:00"
        self.write_candidate(changed)
        with self.assertRaisesRegex(ValueError, "mixed timezone"):
            self.compare()

    def test_constant_and_single_row_targets_have_null_r2(self):
        for frame in (self.frame.assign(y_true=0, y_pred=1), self.frame.iloc[:1]):
            with self.subTest(rows=len(frame)):
                frame.to_csv(self.baseline, index=False)
                self.write_candidate(frame)
                report = self.compare()
                self.assertIsNone(report["overall"]["baseline"]["r2"])
                self.assertIsNone(report["overall"]["delta"]["r2"])
                json.dumps(report, allow_nan=False)

    def test_missing_columns_empty_files_and_duplicate_headers_are_rejected(self):
        self.write_candidate(self.frame.drop(columns="plant_id"))
        with self.assertRaisesRegex(ValueError, "missing required columns"):
            self.compare()
        self.write_candidate(self.frame.iloc[:0])
        with self.assertRaisesRegex(ValueError, "no rows"):
            self.compare()
        self.candidate.write_text("plant_id,timestamp,y_true,y_pred,y_pred\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "duplicate CSV column"):
            self.compare()

    def test_empty_identifiers_are_rejected(self):
        changed = self.frame.copy()
        changed.loc[0, "plant_id"] = " "
        self.write_candidate(changed)
        with self.assertRaisesRegex(ValueError, "empty identifier"):
            self.compare()


if __name__ == "__main__":
    unittest.main()
