"""Projection integrity tests; fixture scores are not PV accuracy evidence."""
from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
import shutil
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest

import numpy as np
import pandas as pd

from solar_forecast.evaluation.model_selection import BenchmarkModelSelector
from solar_forecast.reporting.benchmark_analytics import BenchmarkAnalyticsService
from solar_forecast.reporting.model_analytics import ModelAnalyticsService


class BenchmarkDashboardTests(unittest.TestCase):
    def setUp(self):
        self.temporary = TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.run = self.root / "artifacts/benchmarks/fixture"
        self.task_dir = self.run / "horizon_3"
        self.provenance = {"dataset_fingerprint": "test-fixture-only", "optimization_scope": "bounded_observed_pilot"}
        self.selection = BenchmarkModelSelector(self.task_dir, selection_gap_hours=3).run(
            self.predictions("2025-01-01", 48), self.predictions("2025-01-06", 12),
            evaluation_contract={"horizon_hours": 3, "task": "historical_forecast", "information_set": "observed_measurements_through_forecast_origin"},
            provenance=self.provenance,
        )
        self.manifest = {
            "contract": "solar-optimized-benchmark.v1", "status": "completed", "execution_mode": "full",
            "provenance": self.provenance,
            "tasks": [{"horizon_hours": 3, "selection_path": "horizon_3/selection.json",
                       "optimization": {"xgboost": {"validation_mae": 0.5, "feature_columns": ["lag"], "run_dir": "/private/path", "parameters": {"max_depth": 3}}}}],
        }
        self.selection["files"] = {name: Path(path).name for name, path in self.selection["files"].items()}
        self.write_selection()
        self.write_manifest()

    @staticmethod
    def predictions(start, periods):
        timestamps = pd.date_range(start, periods=periods, freq="h")
        actual = np.arange(periods) % 7 + 1.0
        first = pd.DataFrame({
            "timestamp": timestamps, "forecast_origin": timestamps - pd.Timedelta(hours=3),
            "horizon_hours": 3, "plant_id": "p1", "plant": "발전소1", "region": "지역1",
            "y_true": actual, "xgb_pred": actual + 1, "cnn_pred": actual - 1, "persistence_pred": actual + 2,
        })
        second = first.assign(plant_id="p2", plant="발전소2", xgb_pred=actual - 1, cnn_pred=actual + 1)
        return pd.concat([first, second], ignore_index=True)

    def write_manifest(self):
        (self.run / "manifest.json").write_text(json.dumps(self.manifest), encoding="utf-8")

    def write_selection(self):
        (self.task_dir / "selection.json").write_text(json.dumps(self.selection), encoding="utf-8")

    def result(self):
        return BenchmarkAnalyticsService(self.root).build()

    def prediction_path(self):
        return self.task_dir / self.selection["files"]["test_predictions"]

    def test_completed_artifact_projects_selection_test_series_and_scope(self):
        result = self.result()
        self.assertEqual(result["status"], "ready")
        task = result["tasks"][0]
        self.assertEqual(task["coverage"]["test_rows"], 24)
        self.assertEqual(task["coverage"]["plants"], 2)
        self.assertEqual(len(task["series"]), 24)
        self.assertEqual(task["provenance"]["optimization_scope"], "bounded_observed_pilot")
        models = {row["id"]: row for row in task["models"]}
        self.assertEqual(set(models), {"xgboost", "cnn_bilstm", "hybrid", "persistence"})
        self.assertEqual(models["xgboost"]["test_metrics"]["pooled"]["mae"], 1)
        self.assertEqual(models["xgboost"]["test_metrics"]["national"]["mae"], 0)
        self.assertNotIn("run_dir", task["optimization"]["xgboost"])
        self.assertNotIn(str(self.root), json.dumps(result))

    def test_smoke_and_incomplete_runs_are_excluded(self):
        for field, value in (("execution_mode", "smoke"), ("status", "running")):
            with self.subTest(field=field):
                self.manifest.update(status="completed", execution_mode="full")
                self.manifest[field] = value
                self.write_manifest()
                result = self.result()
                self.assertEqual(result["status"], "empty")
                self.assertEqual(result["excluded_runs"]["non_full_or_incomplete"], 1)

    def test_changed_prediction_file_is_excluded(self):
        path = self.prediction_path()
        path.write_bytes(path.read_bytes() + b"changed")
        self.assertEqual(self.result()["status"], "empty")

    def test_dataset_horizon_and_task_contracts_must_match(self):
        original = json.loads(json.dumps(self.selection))
        for change in ("dataset", "horizon", "task"):
            with self.subTest(change=change):
                self.selection = json.loads(json.dumps(original))
                if change == "dataset":
                    self.selection["provenance"]["dataset_fingerprint"] = "different"
                elif change == "horizon":
                    self.selection["evaluation_contract"]["horizon_hours"] = 24
                else:
                    self.selection["evaluation_contract"]["task"] = "observed_conditions_estimation"
                self.write_selection()
                self.assertEqual(self.result()["status"], "empty")

    def test_wrong_origin_and_selected_predictions_rejected_even_with_updated_hash(self):
        path = self.prediction_path()
        original = pd.read_csv(path)
        for change in ("origin", "selected"):
            with self.subTest(change=change):
                frame = original.copy()
                if change == "origin":
                    frame.loc[0, "forecast_origin"] = frame.loc[0, "timestamp"]
                else:
                    frame.loc[0, "selected_pred"] += 100
                frame.to_csv(path, index=False)
                self.selection["files_sha256"]["test_predictions"] = sha256(path.read_bytes()).hexdigest()
                self.write_selection()
                self.assertEqual(self.result()["status"], "empty")

    def test_misreported_final_metrics_are_rejected(self):
        self.selection["test_metrics"]["xgboost"]["pooled"]["mae"] = 0
        self.write_selection()
        self.assertEqual(self.result()["status"], "empty")

    def test_selection_path_cannot_escape_run_directory(self):
        self.manifest["tasks"][0]["selection_path"] = "../outside/selection.json"
        self.write_manifest()
        self.assertEqual(self.result()["status"], "empty")

    def test_complete_benchmark_can_be_relocated_without_rewriting_artifacts(self):
        original = self.result()
        with TemporaryDirectory() as copied:
            destination = Path(copied)
            shutil.copytree(self.root / "artifacts", destination / "artifacts")
            relocated = BenchmarkAnalyticsService(destination).build()
        self.assertEqual(relocated, original)
        self.assertEqual(relocated["status"], "ready")

    def test_overlapping_selection_and_test_periods_are_rejected(self):
        self.selection["periods"]["selection"]["end"] = self.selection["periods"]["test"]["start"]
        self.write_selection()
        self.assertEqual(self.result()["status"], "empty")

    def test_legacy_comparisons_require_matching_task_and_information(self):
        contract = {"dataset_fingerprint": "fixture", "target": "generation_mwh", "target_unit": "MWh", "horizon_hours": 3,
                    "test_start": "2025-01-01", "test_end": "2025-01-02", "prediction_key": ["timestamp", "plant_id"]}
        signature = lambda c: ModelAnalyticsService._comparison_signature(SimpleNamespace(details={"evaluation_contract": c}))
        self.assertIsNone(signature(contract))
        contract.update(task="historical_forecast", information_set="observed_measurements_through_forecast_origin")
        self.assertIsNotNone(signature(contract))
        self.assertNotEqual(signature(contract), signature({**contract, "task": "observed_conditions_estimation"}))
        self.assertNotEqual(signature(contract), signature({**contract, "information_set": "observed_at_target"}))


if __name__ == "__main__":
    unittest.main()
