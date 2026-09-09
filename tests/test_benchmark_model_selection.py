"""Decision-contract tests use tiny fixtures, not real-model accuracy evidence."""
from __future__ import annotations

import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import numpy as np
import pandas as pd

from solar_forecast.evaluation.model_selection import BenchmarkModelSelector
from solar_forecast.infrastructure.artifact_store import sha256_file
from solar_forecast.models.hybrid.dynamic_gate import ExplainableDynamicGate


def predictions(start: str, periods: int = 48, *, horizon: int = 3) -> pd.DataFrame:
    timestamps = pd.date_range(start, periods=periods, freq="h")
    truth = np.arange(periods) % 7 + 2.0
    return pd.DataFrame({
        "timestamp": timestamps, "forecast_origin": timestamps - pd.Timedelta(hours=horizon),
        "horizon_hours": horizon, "plant_id": "operator:p1", "plant": "발전소1", "region": "지역1",
        "y_true": truth, "xgb_pred": truth + 1.0, "cnn_pred": truth - 1.0,
        "persistence_pred": truth + 2.0,
    })


def artifact_path(report: dict, name: str) -> Path:
    return Path(report["selection_path"]).parent / report["files"][name]


class BenchmarkModelSelectionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.calibration = predictions("2025-01-01")
        self.test = predictions("2025-01-05", periods=12)
        self.contract = {"horizon_hours": 3, "target": "generation_mwh", "information_set": "observed_until_origin"}
        self.provenance = {"dataset_fingerprint": "fixture-only-not-a-performance-result"}

    def run_selector(self, calibration=None, test=None, *, name="run", margin=0.0):
        return BenchmarkModelSelector(self.root / name, margin, selection_gap_hours=3).run(
            self.calibration if calibration is None else calibration,
            self.test if test is None else test,
            evaluation_contract=self.contract, provenance=self.provenance,
        )

    def test_hybrid_is_selected_on_later_calibration_and_test_cannot_reverse_it(self):
        test = self.test.copy()
        test["xgb_pred"] = test.y_true
        test["cnn_pred"] = test.y_true + 100
        result = self.run_selector(test=test)
        self.assertEqual(result["selected_model"], "hybrid")
        self.assertEqual(result["test_metrics"]["xgboost"]["pooled"]["mae"], 0)
        self.assertGreater(result["test_metrics"]["hybrid"]["pooled"]["mae"], 40)
        self.assertFalse(result["selection_rule"]["test_used_for_selection"])
        self.assertFalse(result["persistence_comparison"]["test"]["beats_persistence"])

    def test_test_truth_changes_only_report_not_gate_or_selection(self):
        first = self.run_selector(name="first")
        altered = self.test.copy()
        altered["y_true"] = altered.y_true * 100
        second = self.run_selector(test=altered, name="second")
        self.assertEqual(first["selected_model"], second["selected_model"])
        self.assertEqual(first["selection_metrics"], second["selection_metrics"])
        pd.testing.assert_frame_equal(pd.read_csv(artifact_path(first, "gate_profiles")), pd.read_csv(artifact_path(second, "gate_profiles")))
        self.assertNotEqual(first["test_metrics"], second["test_metrics"])

    def test_gate_fit_excludes_selection_and_purged_hours(self):
        fit_frames = []
        original = ExplainableDynamicGate.fit

        def record_fit(gate, frame):
            fit_frames.append(frame.copy())
            return original(gate, frame)

        with patch.object(ExplainableDynamicGate, "fit", record_fit):
            result = self.run_selector()
        self.assertEqual(len(fit_frames), 1)
        self.assertEqual(len(fit_frames[0]), 21)
        self.assertEqual(result["purge"]["removed_calibration_rows"], 3)
        self.assertLess(fit_frames[0].timestamp.max(), pd.Timestamp(result["periods"]["selection"]["forecast_origin_start"]))
        self.assertTrue(fit_frames[0].plant.eq("operator:p1").all())

    def test_ties_keep_base_model_and_persistence_is_not_a_champion(self):
        frame = self.calibration.copy()
        frame["cnn_pred"] = frame.xgb_pred
        frame["persistence_pred"] = frame.y_true
        result = self.run_selector(calibration=frame)
        self.assertEqual(result["selected_model"], "xgboost")
        self.assertFalse(result["persistence_comparison"]["selection"]["beats_persistence"])

    def test_minimum_improvement_is_required(self):
        frame = self.calibration.copy()
        later = frame.timestamp >= frame.timestamp.iloc[24]
        frame.loc[later, "xgb_pred"] = frame.loc[later, "y_true"] + 0.1
        frame.loc[later, "cnn_pred"] = frame.loc[later, "y_true"] - 0.08
        result = self.run_selector(calibration=frame, margin=0.9)
        self.assertEqual(result["selected_model"], "cnn_bilstm")
        self.assertLess(result["selection_rule"]["hybrid_relative_improvement"], 0.9)

    def test_group_metrics_sum_generation_and_do_not_clip_negative_r2(self):
        calibration2 = self.calibration.assign(plant_id="operator:p2", plant="발전소2")
        test2 = self.test.assign(plant_id="operator:p2", plant="발전소2", xgb_pred=self.test.y_true - 1)
        result = self.run_selector(calibration=pd.concat([self.calibration, calibration2]), test=pd.concat([self.test, test2]))
        metrics = result["test_metrics"]["xgboost"]
        self.assertEqual(metrics["pooled"]["mae"], 1)
        self.assertEqual(metrics["national"]["mae"], 0)
        self.assertEqual(metrics["region"][0]["mae"], 0)
        self.assertEqual(metrics["national"]["n_samples"], 12)
        bad_test = self.test.assign(xgb_pred=self.test.y_true + 100)
        bad = self.run_selector(test=bad_test, name="negative-r2")
        self.assertLess(bad["test_metrics"]["xgboost"]["pooled"]["r2"], 0)

    def test_constant_target_r2_is_json_null(self):
        test = self.test.assign(y_true=1.0)
        result = self.run_selector(test=test)
        self.assertIsNone(result["test_metrics"]["xgboost"]["pooled"]["r2"])
        json.dumps(result, allow_nan=False)

    def test_distinct_registry_ids_can_share_a_display_name(self):
        calibration2 = self.calibration.assign(plant_id="other-operator:p1")
        test2 = self.test.assign(plant_id="other-operator:p1")
        result = self.run_selector(calibration=pd.concat([self.calibration, calibration2]), test=pd.concat([self.test, test2]))
        self.assertEqual(len(result["test_metrics"]["xgboost"]["plant"]), 2)
        profiles = pd.read_csv(artifact_path(result, "gate_profiles"))
        self.assertEqual(set(profiles.loc[profiles.scope.eq("plant"), "plant"]), {"operator:p1", "other-operator:p1"})

    def test_output_contains_frozen_predictions_hashes_and_exact_provenance(self):
        source = self.calibration.copy(deep=True)
        result = self.run_selector()
        saved = json.loads(Path(result["selection_path"]).read_text())
        self.assertEqual(saved["provenance"], self.provenance)
        self.assertEqual(saved["evaluation_contract"], self.contract)
        for key, filename in saved["files"].items():
            self.assertFalse(Path(filename).is_absolute())
            self.assertEqual(saved["files_sha256"][key], sha256_file(artifact_path(result, key)))
        output = pd.read_csv(artifact_path(result, "test_predictions"))
        self.assertTrue(output.selected_model.eq(result["selected_model"]).all())
        np.testing.assert_allclose(output.selected_pred, output.hybrid_pred)
        self.assertTrue(output.plant.eq("발전소1").all())
        pd.testing.assert_frame_equal(self.calibration, source)
        self.assertFalse(list(self.root.rglob("*.tmp")))

    def test_duplicate_or_nonfinite_or_missing_predictions_are_rejected(self):
        invalid = {
            "duplicate": pd.concat([self.test, self.test.iloc[:1]]),
            "nonfinite": self.test.assign(xgb_pred=np.inf),
            "missing": self.test.drop(columns=["forecast_origin"]),
            "missing_time": self.test.assign(timestamp=pd.NaT),
            "empty": self.test.iloc[:0],
        }
        for name, frame in invalid.items():
            with self.subTest(name=name), self.assertRaises(ValueError):
                self.run_selector(test=frame, name=name)
            self.assertFalse((self.root / name / "selection.json").exists())

    def test_horizon_and_truth_mismatches_are_rejected(self):
        invalid = {
            "horizon": self.test.assign(horizon_hours=24),
            "origin": self.test.assign(forecast_origin=self.test.forecast_origin + pd.Timedelta(hours=1)),
            "truth": self.test.assign(y_true_cnn=self.test.y_true + 1),
            "identity": self.test.assign(region="different-region"),
        }
        for name, frame in invalid.items():
            with self.subTest(name=name), self.assertRaises(ValueError):
                self.run_selector(test=frame, name=name)

    def test_calibration_and_test_origin_overlap_is_rejected(self):
        test = predictions("2025-01-02 23:00", periods=12)
        with self.assertRaisesRegex(ValueError, "precede every Test forecast origin"):
            self.run_selector(test=test)

    def test_short_calibration_cannot_skip_the_purge(self):
        with self.assertRaisesRegex(ValueError, "too short"):
            self.run_selector(calibration=self.calibration.iloc[:6])

    def test_existing_evidence_is_not_overwritten(self):
        result = self.run_selector()
        original = Path(result["selection_path"]).read_bytes()
        with self.assertRaises(FileExistsError):
            self.run_selector()
        self.assertEqual(Path(result["selection_path"]).read_bytes(), original)


if __name__ == "__main__":
    unittest.main()
