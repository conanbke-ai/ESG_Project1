"""Benchmark restarts reuse training state without reusing changed experiments."""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import tempfile
import unittest

from solar_forecast.config_loader import ModelJobConfig, PROJECT_ROOT
from solar_forecast.evaluation.experiment_config import build_candidate_configs, load_experiment_config
from solar_forecast.models.cnn_bilstm.input_preprocessing import INPUT_PREPROCESSING_CONTRACT
from solar_forecast.models.shared.checkpoint_store import TrainingCheckpointStore, training_fingerprint


class BenchmarkCheckpointIdentityTests(unittest.TestCase):
    def _experiment(self, root: Path) -> dict:
        source = root / "observed.csv"
        source.write_text("generation_mwh\n1.0\n", encoding="utf-8")
        values = load_experiment_config(PROJECT_ROOT / "config/experiments/observed_calendar_candidate.json")
        values["input_dataset"] = str(source)
        return values

    @staticmethod
    def _with_values(config: ModelJobConfig, values: dict) -> ModelJobConfig:
        return ModelJobConfig(config.model, config.profile, values, config.source)

    def test_new_report_directory_reuses_every_candidate_checkpoint(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            values = self._experiment(root)
            seen = set()
            for horizon in values["horizons_hours"]:
                first = build_candidate_configs(values, horizon, root / "first_run")
                restarted = build_candidate_configs(values, horizon, root / "restarted_run")
                for (name, before), (other_name, after) in zip(first, restarted):
                    with self.subTest(model=before.model, candidate=name, horizon=horizon):
                        self.assertEqual(name, other_name)
                        self.assertNotEqual(before.values["output_root"], after.values["output_root"])
                        old_store = TrainingCheckpointStore.from_config(before)
                        new_store = TrainingCheckpointStore.from_config(after)
                        self.assertEqual(old_store.directory, new_store.directory)
                        self.assertTrue(new_store.resume)
                        self.assertNotIn(new_store.fingerprint, seen)
                        seen.add(new_store.fingerprint)
                        if before.model == "cnn_bilstm":
                            self.assertEqual(before.values["input_preprocessing_contract"], INPUT_PREPROCESSING_CONTRACT)
            self.assertEqual(len(seen), 18)

    def test_changed_training_meaning_never_reuses_state(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            candidates = build_candidate_configs(self._experiment(root), 24, root / "run")
            for model in ("xgboost", "cnn_bilstm"):
                config = next(config for _, config in candidates if config.model == model)
                fingerprint = training_fingerprint(config)
                changes = {
                    "seed": 43,
                    "forecast_horizon_hours": 72,
                    "feature_columns": list(reversed(config.values["feature_columns"])),
                    "train_end": "2024-05-31T23:00:00",
                    "quality_filter_column": "different_eligibility_policy",
                    "input_preprocessing_contract": "different_preprocessing.v3",
                }
                if model == "cnn_bilstm":
                    changes["sequence_length"] = 336
                for key, value in changes.items():
                    with self.subTest(model=model, field=key):
                        changed = deepcopy(config.values)
                        changed[key] = value
                        self.assertNotEqual(fingerprint, training_fingerprint(self._with_values(config, changed)))
                changed = deepcopy(config.values)
                changed["optimizer"]["search_space"] = {"different_search": {"type": "fixed", "value": 1}}
                self.assertNotEqual(fingerprint, training_fingerprint(self._with_values(config, changed)))
                source = Path(config.values["input_dataset"])
                previous = source.read_text(encoding="utf-8")
                source.write_text(previous + "2.0\n", encoding="utf-8")
                self.assertNotEqual(fingerprint, training_fingerprint(config))
                source.write_text(previous, encoding="utf-8")

    def test_trial_budget_extension_keeps_same_training_identity(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            values = self._experiment(root)
            first = build_candidate_configs(values, 1, root / "first")
            for settings in values["models"].values():
                settings["max_trials_per_candidate"] += 1
                settings["timeout_seconds_per_candidate"] += 60
            second = build_candidate_configs(values, 1, root / "second")
            for (_, before), (_, after) in zip(first, second):
                self.assertEqual(training_fingerprint(before), training_fingerprint(after))

    def test_standalone_training_keeps_existing_output_path_identity(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = build_candidate_configs(self._experiment(root), 1, root / "first")[0][1]
            original = deepcopy(config.values)
            original.pop("checkpoint_identity_contract")
            changed = deepcopy(original)
            changed["output_root"] = str(root / "second")
            self.assertNotEqual(
                training_fingerprint(self._with_values(config, original)),
                training_fingerprint(self._with_values(config, changed)),
            )


if __name__ == "__main__":
    unittest.main()
