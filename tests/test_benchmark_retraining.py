"""Reject real retraining drift while allowing timestamp/output relocation."""
from __future__ import annotations

from copy import deepcopy
import csv
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

from tools.verify_benchmark_retraining import (
    compare_benchmarks, compare_predictions, fresh_experiment, semantic_config,
)


def write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value), encoding="utf-8")


def write_predictions(path: Path, *, prediction: float = 0.5, truth: float = 0.6,
                      persistence: float = 0.4, plant: str = "plant_a", final: bool = False) -> None:
    row = {"plant_id": plant, "timestamp": "2024-12-01 12:00:00",
           "forecast_origin": "2024-12-01 11:00:00", "horizon_hours": 1,
           "plant": "A", "region": "R", "y_true": truth, "persistence_pred": persistence}
    row.update({name: prediction for name in (("xgb_pred", "cnn_pred", "hybrid_pred", "selected_pred") if final else ("y_pred",))})
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(row))
        writer.writeheader()
        writer.writerow(row)


def benchmark_fixture(root: Path, timestamp: str) -> None:
    specs = [("xgboost", "weather", None), ("cnn_bilstm", "weather_lookback_24h", 24),
             ("cnn_bilstm", "weather_lookback_168h", 168)]
    experiment = {"input_dataset": "/observed/input.csv", "output_root": str(root), "seed": 42,
                  "models": {name: {"training_overrides": {"checkpoint": {"root": str(root / "checkpoints"), "resume": False}},
                                    "optimizer_overrides": {"storage_path": str(root / "trials.db")}}
                             for name in ("xgboost", "cnn_bilstm")}}
    selected = {}
    for model, candidate_id, lookback in specs:
        relative = Path("candidates/horizon_1h") / model / candidate_id / timestamp
        candidate = root / relative
        config = {"model": model, "benchmark_candidate_id": candidate_id, "forecast_horizon_hours": 1,
                  "sequence_length": lookback, "seed": 42, "feature_columns": ["temperature_c"],
                  "input_dataset": "/observed/input.csv", "output_root": str(candidate),
                  "checkpoint": {"root": str(root / "checkpoints"), "resume": False},
                  "optimizer": {"storage_path": str(root / "trials.db"), "max_trials": 1}}
        details = {"evaluation_contract": {"dataset_fingerprint": "input-sha", "horizon_hours": 1},
                   "temporal_split": {"train_end": "2024-01-01"},
                   "checkpoint": {"resumed": False},
                   "optimizer": {"summary_path": str(candidate / "optimization_summary.json"), "best_params": {"depth": 3}}}
        for split in ("validation", "calibration", "test"):
            path = candidate / f"{split}_predictions.csv"
            write_predictions(path)
            details[f"{split}_predictions"] = str(path)
        write_json(candidate / "optimization_summary.json", {"existing_finished_trials": 0, "executed_trials": 1})
        write_json(candidate / "resolved_config.json", config)
        write_json(candidate / "manifest.json", {"status": "completed", "model": model, "details": details})
        if model == "xgboost" or lookback == 24:
            selected[model] = {"candidate_id": candidate_id, "sequence_length": lookback,
                               "feature_columns": ["temperature_c"], "parameters": {"depth": 3},
                               "run_dir": relative.as_posix(), "resolved_config": (relative / "resolved_config.json").as_posix()}
    write_json(root / "experiment.json", experiment)
    write_json(root / "plan.json", {"tasks": [{"horizon_hours": 1, "candidates": [
        {"model": model, "candidate_id": candidate} for model, candidate, _ in specs]}]})
    write_json(root / "horizon_1h/base_selection.json", {"test_used_for_selection": False, "selected": selected})
    write_json(root / "horizon_1h/selection.json", {
        "selected_model": "xgboost", "selection_rule": {"test_used_for_selection": False},
        "files": {"test_predictions": "test_predictions.csv"},
    })
    write_predictions(root / "horizon_1h/test_predictions.csv", final=True)
    write_json(root / "manifest.json", {
        "status": "completed", "execution_mode": "full",
        "provenance": {"dataset_fingerprint": "input-sha", "source_sha256": "source-sha",
                       "runtime": {"torch": "pinned"}, "optimization_scope": "bounded_observed_pilot"},
        "tasks": [{"horizon_hours": 1, "selected_model": "xgboost", "optimization": selected,
                   "selection_path": "horizon_1h/selection.json"}],
    })


class BenchmarkRetrainingTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.reference, self.repeated = self.root / "reference", self.root / "repeated"
        benchmark_fixture(self.reference, "first_timestamp")
        benchmark_fixture(self.repeated, "second_timestamp")

    def candidate_file(self, name: str = "test_predictions.csv", *, selected: bool = False) -> Path:
        candidate = "weather_lookback_24h" if selected else "weather_lookback_168h"
        return self.repeated / "candidates/horizon_1h/cnn_bilstm" / candidate / "second_timestamp" / name

    def change_json(self, path: Path, change) -> None:
        value = json.loads(path.read_text())
        change(value)
        write_json(path, value)

    def test_all_candidates_and_splits_compared_across_timestamp_directories(self):
        report = compare_benchmarks(self.reference, self.repeated)
        self.assertEqual(report["status"], "passed")
        self.assertEqual(len(report["candidates"]), 3)
        self.assertTrue(all(set(item["splits"]) == {"validation", "calibration", "test"} for item in report["candidates"]))

    def test_relocated_artifacts_use_new_folder_even_when_original_exists(self):
        relocated = self.root / "relocated" / self.repeated.name
        shutil.copytree(self.repeated, relocated)
        write_predictions(self.candidate_file(), prediction=0.9)
        self.assertEqual(compare_benchmarks(self.reference, relocated)["status"], "passed")

    def test_unselected_candidate_drift_fails_for_each_split(self):
        for split in ("validation", "calibration", "test"):
            with self.subTest(split=split):
                path = self.candidate_file(f"{split}_predictions.csv")
                write_predictions(path, prediction=0.50001)
                with self.assertRaisesRegex(ValueError, "Prediction drift"):
                    compare_benchmarks(self.reference, self.repeated)
                write_predictions(path)

    def test_small_prediction_delta_passes_but_truth_must_be_exact(self):
        reference = self.root / "before.csv"
        repeated = self.root / "after.csv"
        write_predictions(reference)
        write_predictions(repeated, prediction=0.5000001)
        self.assertEqual(compare_predictions(reference, repeated)["rows"], 1)
        for changed in ({"truth": 0.60000001}, {"persistence": 0.40000001}, {"plant": "plant_b"}):
            with self.subTest(changed=changed):
                write_predictions(repeated, **changed)
                with self.assertRaises(ValueError):
                    compare_predictions(reference, repeated)

    def test_duplicate_forecast_keys_fail(self):
        path = self.candidate_file()
        with path.open("a", encoding="utf-8") as stream:
            stream.write(path.read_text().splitlines()[1] + "\n")
        with self.assertRaisesRegex(ValueError, "Duplicate"):
            compare_benchmarks(self.reference, self.repeated)

    def test_existing_output_is_rejected_without_overwriting_evidence(self):
        output = self.root / "existing"
        output.mkdir()
        evidence = output / "retraining_report.json"
        evidence.write_text("existing evidence")
        script = Path(__file__).resolve().parents[1] / "tools/verify_benchmark_retraining.py"
        result = subprocess.run([sys.executable, str(script), "--output-dir", str(output)],
                                capture_output=True, text=True, check=False)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("existing evidence is never overwritten", result.stderr)
        self.assertEqual(evidence.read_text(), "existing evidence")

    def test_frozen_model_change_fails_even_with_identical_predictions(self):
        self.change_json(self.repeated / "manifest.json", lambda v: v["tasks"][0].update(selected_model="hybrid"))
        self.change_json(self.repeated / "horizon_1h/selection.json", lambda v: v.update(selected_model="hybrid"))
        with self.assertRaisesRegex(ValueError, "Frozen selected"):
            compare_benchmarks(self.reference, self.repeated)

    def test_frozen_lookback_change_fails(self):
        def change(values):
            values["cnn_bilstm"].update(candidate_id="weather_lookback_168h", sequence_length=168)
        self.change_json(self.repeated / "manifest.json", lambda v: change(v["tasks"][0]["optimization"]))
        self.change_json(self.repeated / "horizon_1h/base_selection.json", lambda v: change(v["selected"]))
        with self.assertRaisesRegex(ValueError, "Frozen selected"):
            compare_benchmarks(self.reference, self.repeated)

    def test_resolved_training_setting_change_is_not_treated_as_relocation(self):
        self.change_json(self.candidate_file("resolved_config.json"), lambda v: v.update(seed=43))
        with self.assertRaisesRegex(ValueError, "Resolved model configuration"):
            compare_benchmarks(self.reference, self.repeated)

    def test_final_blend_drift_fails_after_identical_candidate_predictions(self):
        write_predictions(self.repeated / "horizon_1h/test_predictions.csv", prediction=0.6, final=True)
        with self.assertRaisesRegex(ValueError, "Prediction drift"):
            compare_benchmarks(self.reference, self.repeated)

    def test_missing_candidate_rejected_instead_of_comparing_intersection(self):
        self.candidate_file("resolved_config.json").unlink()
        with self.assertRaisesRegex(ValueError, "complete experiment plan"):
            compare_benchmarks(self.reference, self.repeated)

    def test_source_or_input_fingerprint_change_fails(self):
        for field in ("source_sha256", "dataset_fingerprint", "runtime"):
            with self.subTest(field=field):
                path = self.repeated / "manifest.json"
                original = path.read_text()
                self.change_json(path, lambda v: v["provenance"].update({field: "different"}))
                with self.assertRaisesRegex(ValueError, "provenance"):
                    compare_benchmarks(self.reference, self.repeated)
                path.write_text(original)

    def test_existing_optimizer_trials_or_resumed_checkpoint_fail(self):
        summary = self.candidate_file("optimization_summary.json")
        self.change_json(summary, lambda v: v.update(existing_finished_trials=1))
        with self.assertRaisesRegex(ValueError, "fresh optimizer"):
            compare_benchmarks(self.reference, self.repeated)
        self.change_json(summary, lambda v: v.update(existing_finished_trials=0))
        self.change_json(self.candidate_file("manifest.json"), lambda v: v["details"]["checkpoint"].update(resumed=True))
        with self.assertRaisesRegex(ValueError, "reused a checkpoint"):
            compare_benchmarks(self.reference, self.repeated)

    def test_fresh_configuration_changes_only_storage_locations(self):
        values = json.loads((self.reference / "experiment.json").read_text())
        before = deepcopy(values)
        cloned = fresh_experiment(values, self.root / "input.csv", self.root / "isolated")
        self.assertEqual(values, before)
        self.assertEqual(semantic_config(values, experiment=True), semantic_config(cloned, experiment=True))
        for settings in cloned["models"].values():
            self.assertFalse(settings["training_overrides"]["checkpoint"]["resume"])
            self.assertIn("isolated", settings["optimizer_overrides"]["storage_path"])
        cloned["seed"] = 43
        self.assertNotEqual(semantic_config(values, experiment=True), semantic_config(cloned, experiment=True))


if __name__ == "__main__":
    unittest.main()
