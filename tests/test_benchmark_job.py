"""Benchmark orchestration tests use labelled fixtures, never accuracy evidence."""
from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock

import numpy as np
import pandas as pd

from solar_forecast.config_loader import PROJECT_ROOT
from solar_forecast.evaluation.experiment_config import load_experiment_config
from solar_forecast.evaluation.forecast_samples import forecast_evaluation_contract
from solar_forecast.infrastructure.artifact_store import write_json_atomic, write_manifest
from solar_forecast.jobs.benchmark_job import BenchmarkService, align_prediction_frames, _validate_search_cohorts
from solar_forecast.models.shared.checkpoint_store import dataset_signature


def _predictions(start: str, rows: int, error: float) -> pd.DataFrame:
    timestamps = pd.date_range(start, periods=rows, freq="h")
    actual = 2 + np.sin(np.arange(rows) / 7)
    return pd.DataFrame({"timestamp": timestamps, "forecast_origin": timestamps - pd.Timedelta(hours=1),
        "horizon_hours": 1, "plant_id": "fixture_plant", "plant": "fixture", "region": "fixture_region",
        "y_true": actual, "y_pred": actual + error, "persistence_pred": actual + 1})


class _FixtureTrainingService:
    def run(self, config, *, smoke=False):
        run_dir = Path(config.values["output_root"]) / "fixture_run"
        run_dir.mkdir(parents=True)
        candidate = config.values["benchmark_candidate_id"]
        good = candidate.startswith("observed_weather_history") and config.values.get("sequence_length", 24) == 24
        error = 0.1 if good else 0.8
        if config.model == "cnn_bilstm":
            error *= -1
        paths = {}
        for split, start, rows in (("validation", "2024-01-01", 60), ("calibration", "2024-01-10", 200), ("test", "2024-02-01", 60)):
            # Deliberately reverse Test rankings. Validation must still win.
            split_error = (2 if good else 0.01) if split == "test" else error
            frame = _predictions(start, rows, split_error)
            path = run_dir / f"{split}_predictions.csv"
            frame.to_csv(path, index=False)
            paths[f"{split}_predictions"] = str(path)
        contract = {"dataset_fingerprint": dataset_signature(Path(config.values["input_dataset"])),
                    "target": "generation_mwh", "target_unit": "MWh",
                    **forecast_evaluation_contract("historical_forecast", 1, legacy_task="unused")}
        write_manifest(run_dir / "manifest.json", status="completed", model=config.model,
                       run_id="fixture_run", details={**paths, "evaluation_contract": contract})
        return run_dir


class BenchmarkJobTests(unittest.TestCase):
    def test_base_selection_ignores_reversed_test_ranking(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "fixture.csv"
            source.write_text("fixture_only\n1\n")
            config = load_experiment_config(PROJECT_ROOT / "config/experiments/optimized.json")
            config.update(input_dataset=str(source), output_root=str(root / "benchmarks"),
                          horizons_hours=[1], selection_gap_hours=1,
                          optimization_scope="test_fixture")
            path = root / "experiment.json"
            write_json_atomic(path, config)
            run_dir = BenchmarkService(training_service=_FixtureTrainingService()).run(path)
            manifest = json.loads((run_dir / "manifest.json").read_text())
            self.assertEqual(manifest["status"], "completed")
            selection = json.loads((run_dir / "horizon_1h/base_selection.json").read_text())
            self.assertFalse(selection["test_used_for_selection"])
            self.assertEqual(selection["selected"]["xgboost"]["candidate_id"], "observed_weather_history")
            self.assertEqual(selection["selected"]["cnn_bilstm"]["candidate_id"], "observed_weather_history_lookback_24h")
            result = json.loads((run_dir / "horizon_1h/selection.json").read_text())
            self.assertEqual(result["selected_model"], "hybrid")

    def test_alignment_rejects_truth_change_and_low_common_coverage(self):
        first = _predictions("2024-01-01", 20, 0.1)
        second = first.copy()
        second.loc[0, "y_true"] += 1
        with self.assertRaisesRegex(ValueError, "inconsistent y_true"):
            align_prediction_frames({"a": first, "b": second}, minimum_coverage=0.95)
        with self.assertRaisesRegex(ValueError, "coverage"):
            align_prediction_frames({"a": first, "b": first.iloc[:5]}, minimum_coverage=0.95)

    def test_alignment_reports_every_dropped_row(self):
        first = _predictions("2024-01-01", 20, 0.1)
        frames, coverage = align_prediction_frames({"a": first, "b": first.iloc[1:]}, minimum_coverage=0.95)
        self.assertEqual(coverage["dropped_rows"], {"a": 1, "b": 0})
        self.assertEqual(len(frames["a"]), 19)

    def test_search_candidate_cohorts_must_be_identical(self):
        cohort = {
            "contract": "solar-validation-cohort.v1",
            "key_columns": [
                "plant_id",
                "timestamp",
                "forecast_origin",
                "horizon_hours",
            ],
            "rows": 100,
            "sha256": "same",
            "target_start": "2024-07-01T00:00:00",
            "target_end": "2024-09-30T23:00:00",
        }
        coverage = _validate_search_cohorts(
            {"xgb": cohort, "cnn": dict(cohort)},
            minimum_coverage=0.95,
        )
        self.assertEqual(coverage["common_rows"], 100)
        self.assertEqual(coverage["common_fraction"], 1.0)
        self.assertEqual(coverage["dropped_rows"], {"xgb": 0, "cnn": 0})

        changed = dict(cohort, sha256="different")
        with self.assertRaisesRegex(ValueError, "cohorts differ"):
            _validate_search_cohorts(
                {"xgb": cohort, "cnn": changed},
                minimum_coverage=0.95,
            )

    def test_search_candidate_cohort_cannot_be_missing(self):
        with self.assertRaisesRegex(ValueError, "omitted"):
            _validate_search_cohorts(
                {"xgb": {}, "cnn": {}},
                minimum_coverage=0.95,
            )

    def test_legacy_experiment_schema_is_not_silently_executed(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "legacy.json"
            path.write_text('{"name":"legacy_24h"}')
            with self.assertRaisesRegex(ValueError, "contract"):
                load_experiment_config(path)

    def test_gzip_partition_directory_reaches_training_but_parquet_does_not(self):
        import gzip

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "parts"
            source.mkdir()
            config = load_experiment_config(PROJECT_ROOT / "config/experiments/optimized.json")
            config.update(input_dataset=str(source), output_root=str(root / "runs"))
            path = root / "experiment.json"
            write_json_atomic(path, config)
            trainer = Mock()
            trainer.run.side_effect = RuntimeError("fixture reached trainer")
            (source / "unsupported.parquet").write_bytes(b"fixture")
            with self.assertRaises(FileNotFoundError):
                BenchmarkService(training_service=trainer).run(path)
            trainer.run.assert_not_called()
            self.assertFalse((root / "runs").exists())
            (source / "unsupported.parquet").unlink()
            with gzip.open(source / "part.csv.gz", "wt") as stream:
                stream.write("fixture_only\n1\n")
            with self.assertRaisesRegex(RuntimeError, "fixture reached trainer"):
                BenchmarkService(training_service=trainer).run(path)
            trainer.run.assert_called_once()


if __name__ == "__main__":
    unittest.main()
