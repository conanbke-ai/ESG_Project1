"""정확한 시각 표본·연속창·학습 입력 통계의 데이터 전용 검증."""
from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import numpy as np
import pandas as pd

from solar_forecast.config_loader import ModelJobConfig
from solar_forecast.evaluation.forecast_readiness import audit_forecast_frame, run_forecast_readiness


class ForecastReadinessTests(unittest.TestCase):
    def frame(self):
        times = pd.date_range("2024-01-01", periods=600, freq="h")
        return pd.DataFrame({"plant_id": "p", "timestamp": times, "generation_mwh": np.arange(600, dtype=np.float32),
                             "f": np.arange(600, dtype=np.float32), "all_missing": np.nan,
                             "constant": 7.0})

    def candidates(self, horizon=24, length=24):
        values = {"prediction_task": "historical_forecast", "forecast_horizon_hours": horizon,
                  "target_column": "generation_mwh", "feature_columns": ["f", "all_missing", "constant"],
                  "entity_column": "plant_id", "timestamp_column": "timestamp", "sequence_length": length,
                  "train_end": "2024-01-13T11:00:00", "validation_end": "2024-01-17T15:00:00",
                  "calibration_end": "2024-01-21T19:00:00", "test_end": "2024-01-25T23:00:00", "purge_gap_hours": 24}
        return [("features", ModelJobConfig(model, "test", dict(values), Path("unused"))) for model in ("xgboost", "cnn_bilstm")]

    def test_exact_origin_window_counts_and_train_only_features(self):
        frame = self.frame()
        frame.loc[frame.index >= 300, "f"] = np.nan
        result = audit_forecast_frame(frame, self.candidates(), minimum_coverage=0.95)
        xgb, cnn = result["candidates"]
        self.assertEqual(xgb["splits"]["train"]["usable_samples"], 276)
        self.assertEqual(cnn["splits"]["train"]["usable_samples"], 253)
        self.assertEqual(cnn["splits"]["train"]["losses"]["unavailable_exact_origin"], 24)
        self.assertEqual(cnn["splits"]["train"]["losses"]["insufficient_history"], 23)
        self.assertEqual(cnn["train_feature_sanity"]["rows"], 276)
        self.assertEqual(cnn["train_feature_sanity"]["features"][0]["missing_rows"], 0)
        self.assertEqual(cnn["train_feature_sanity"]["all_missing_features"], ["all_missing"])
        self.assertEqual(cnn["train_feature_sanity"]["constant_observed_features"], ["constant"])
        self.assertTrue(result["coverage_gate_passed"])
        self.assertEqual(result["common_coverage"]["test"]["common_fraction"], 1)
        future_changed = frame.copy()
        future_changed.loc[future_changed.index >= 300, "f"] = -1e9
        again = audit_forecast_frame(future_changed, self.candidates(), minimum_coverage=0.95)
        self.assertEqual(cnn["train_feature_sanity"], again["candidates"][1]["train_feature_sanity"])

    def test_all_missing_cnn_requires_enabled_missing_indicators(self):
        candidates = self.candidates()
        candidates[1][1].values["append_missing_indicators"] = False
        with self.assertRaisesRegex(ValueError, "require missing indicators"):
            audit_forecast_frame(self.frame(), candidates, minimum_coverage=0.95)
        candidates[1][1].values["append_missing_indicators"] = True
        report = audit_forecast_frame(self.frame(), candidates, minimum_coverage=0.95)
        self.assertIn("all_missing_train_feature:all_missing", report["candidates"][1]["warnings"])

    def test_gap_loss_matches_actual_hourly_windows_and_conserves_rows(self):
        frame = self.frame().drop(index=350)
        result = audit_forecast_frame(frame, self.candidates(), minimum_coverage=0.95)
        cnn = result["candidates"][1]
        self.assertEqual(cnn["splits"]["validation"]["losses"]["unavailable_exact_origin"], 1)
        self.assertEqual(cnn["splits"]["validation"]["losses"]["discontinuous_history"], 23)
        self.assertFalse(result["coverage_gate_passed"])
        for candidate in result["candidates"]:
            for partition in candidate["splits"].values():
                self.assertEqual(partition["eligible_target_rows"], partition["usable_samples"] + sum(partition["losses"].values()))

    def test_non_finite_origin_and_target_do_not_become_samples(self):
        frame = self.frame()
        frame.loc[350, "generation_mwh"] = np.inf
        result = audit_forecast_frame(frame, self.candidates(), minimum_coverage=0.5)
        for candidate in result["candidates"]:
            losses = candidate["splits"]["validation"]["losses"]
            self.assertEqual(losses["non_finite_target"], 1)
            self.assertEqual(losses["non_finite_origin"], 1)

    def test_reject_duplicate_non_hourly_timezone_and_empty_partition(self):
        frame = self.frame()
        cases = [pd.concat([frame, frame.iloc[[0]]], ignore_index=True)]
        nonhourly = frame.copy()
        nonhourly.loc[0, "timestamp"] += pd.Timedelta(minutes=1)
        cases.append(nonhourly)
        zoned = frame.copy()
        zoned["timestamp"] = zoned["timestamp"].dt.tz_localize("Asia/Seoul")
        cases.extend([zoned, frame.iloc[:350]])
        for bad in cases:
            with self.subTest(case=str(bad.dtypes)), self.assertRaises(ValueError):
                audit_forecast_frame(bad, self.candidates(), minimum_coverage=0.95)

    def test_feature_subsets_share_rows_and_preserve_order(self):
        candidates = self.candidates()
        original = candidates[1][1]
        subset = ModelJobConfig("cnn_bilstm", "test", {**original.values, "feature_columns": ["constant", "f"]}, original.source)
        result = audit_forecast_frame(self.frame(), [*candidates, ("subset", subset)], minimum_coverage=0.95)
        self.assertEqual(result["candidates"][1]["splits"], result["candidates"][2]["splits"])
        self.assertEqual(result["candidates"][2]["train_feature_sanity"]["feature_order"], ["constant", "f"])

    def test_file_runner_uses_trainer_solar_quality_filters_and_hashes(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            frame = self.frame().assign(energy_source="solar", quality_train_eligible="yes")
            excluded = frame.iloc[:3].copy().assign(plant_id="excluded", energy_source="wind")
            false_quality = frame.iloc[:3].copy().assign(plant_id="excluded", quality_train_eligible="false")
            source = root / "gold.csv"
            pd.concat([frame, excluded, false_quality], ignore_index=True).to_csv(source, index=False)
            configurations = self.candidates()
            models = {}
            for _, cfg in configurations:
                path = root / f"{cfg.model}.json"
                path.write_text(json.dumps({**cfg.values, "model": cfg.model, "numeric_dtype": "float32"}), encoding="utf-8")
                models[cfg.model] = {"config": str(path), "feature_sets": ["full"], "sequence_lengths": [24],
                                     "max_trials_per_candidate": 1, "timeout_seconds_per_candidate": 1}
            spec = {"contract": "solar-optimized-experiment.v1", "input_dataset": str(source),
                    "horizons_hours": [24], "minimum_common_coverage": 0.95, "models": models,
                    "feature_sets": {"full": {}}, "split": {key: configurations[0][1].values[key] for key in
                        ("train_end", "validation_end", "calibration_end", "test_end", "purge_gap_hours")}}
            config = root / "experiment.json"
            config.write_text(json.dumps(spec), encoding="utf-8")
            report = run_forecast_readiness(config, output_path=root / "audit.json", project_root=root)
            self.assertEqual(report["loading"]["scanned_rows"], 606)
            self.assertEqual(report["loading"]["retained_rows"], 600)
            self.assertFalse(report["prediction_or_training_performed"])
            self.assertEqual(len(report["source_files"][0]["sha256"]), 64)
            self.assertEqual(set(report["model_config_sha256"]), {"xgboost", "cnn_bilstm"})
            self.assertTrue((root / "audit.json").exists())
            with self.assertRaises(ValueError):
                run_forecast_readiness(config, output_path=source, project_root=root)
            for action in ("add", "remove"):
                partition_dir = root / action
                partition_dir.mkdir()
                partition = partition_dir / "part.csv"
                partition.write_bytes(source.read_bytes())

                def mutate_partitions(*args, **kwargs):
                    result = audit_forecast_frame(*args, **kwargs)
                    if action == "add":
                        (partition_dir / "new.csv").write_bytes(source.read_bytes())
                    else:
                        partition.unlink()
                    return result

                with self.subTest(partition_change=action), patch(
                    "solar_forecast.evaluation.forecast_readiness.audit_forecast_frame",
                    side_effect=mutate_partitions,
                ), self.assertRaisesRegex(ValueError, "partition"):
                    run_forecast_readiness(config, data_path=partition_dir, project_root=root)


if __name__ == "__main__":
    unittest.main()
