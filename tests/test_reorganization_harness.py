"""Safety and fixture-contract checks independent of ML dependencies."""
import csv
import importlib.util
from pathlib import Path
import tempfile
import unittest


SPEC = importlib.util.spec_from_file_location("reorganization_harness", Path(__file__).parents[1] / "tools" / "verify_model_reorganization.py")
harness = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(harness)


class ReorganizationHarnessTests(unittest.TestCase):
    def test_source_identity_detects_uncommitted_code_and_configuration_changes(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / "src").mkdir()
            (root / "config" / "models").mkdir(parents=True)
            source = root / "src" / "trainer.py"
            config = root / "config" / "models" / "model.json"
            source.write_text("value = 1\n", encoding="utf-8")
            config.write_text('{"epochs": 2}', encoding="utf-8")
            original = harness.source_identity(root)
            source.write_text("value = 2\n", encoding="utf-8")
            changed_code = harness.source_identity(root)
            self.assertNotEqual(original["source_sha256"], changed_code["source_sha256"])
            config.write_text('{"epochs": 3}', encoding="utf-8")
            self.assertNotEqual(changed_code["source_sha256"], harness.source_identity(root)["source_sha256"])
            self.assertEqual(original["git_dirty"], "unavailable")

    def test_existing_evidence_is_never_overwritten(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary)
            evidence = path / "report.json"
            evidence.write_text("retained", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "never overwritten"):
                harness.create_output_directory(path)
            self.assertEqual(evidence.read_text(encoding="utf-8"), "retained")

    def test_fixture_is_frozen_and_has_disjoint_entities_and_missingness(self):
        with tempfile.TemporaryDirectory() as temporary:
            first, second = Path(temporary) / "first.csv", Path(temporary) / "second.csv"
            features = ["temperature_c", "tilt_deg", "generation_lag_168h_mwh", "is_daylight"]
            harness.create_synthetic_fixture(first, features)
            harness.create_synthetic_fixture(second, features)
            self.assertEqual(harness.sha256_file(first), harness.sha256_file(second))
            with first.open(encoding="utf-8", newline="") as stream:
                rows = list(csv.DictReader(stream))
            self.assertEqual(len(rows), 960)
            self.assertEqual(len({(row["timestamp"], row["plant_id"]) for row in rows}), 960)
            self.assertEqual({row["plant_id"] for row in rows}, {"synthetic-0", "synthetic-1"})
            self.assertTrue(all(row["tilt_deg"] == "" for row in rows))
            self.assertTrue(any(row["temperature_c"] == "" for row in rows))
            self.assertTrue(any(float(row["generation_mwh"]) > 0.5 for row in rows))
            self.assertTrue(any(float(row["generation_mwh"]) == 0 for row in rows))

    def test_small_budget_does_not_modify_original_configuration(self):
        original = {"optimizer": {"enabled": True}, "checkpoint": {"resume": True}, "feature_columns": ["f"], "validation_fraction": 0.15}
        configured = harness.experiment_config(original, Path("/frozen.csv"), Path("/evidence"))
        self.assertTrue(original["optimizer"]["enabled"])
        self.assertTrue(original["checkpoint"]["resume"])
        self.assertFalse(configured["checkpoint"]["enabled"])
        self.assertFalse(configured["checkpoint"]["resume"])
        self.assertEqual(configured["feature_columns"], ["f"])
        self.assertEqual(configured["validation_fraction"], 0.15)


if __name__ == "__main__":
    unittest.main()
