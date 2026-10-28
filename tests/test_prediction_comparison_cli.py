"""Exercise prediction comparison as a real process, including artifact safety."""
from __future__ import annotations

import csv
import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import textwrap
import unittest


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class PredictionComparisonCliTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.baseline = self.root / "baseline.csv"
        self.candidate = self.root / "candidate.csv"
        self.report = self.root / "comparison.json"
        self.write_predictions(self.baseline)
        self.write_predictions(self.candidate)

    def write_predictions(self, path, *, shift=0):
        with path.open("w", encoding="utf-8", newline="") as stream:
            writer = csv.writer(stream)
            writer.writerow(["plant_id", "timestamp", "region", "y_true", "y_pred"])
            writer.writerow(["001", "2025-01-01 00:00", "부산", 0, shift])
            writer.writerow(["001", "2025-01-01 01:00", "부산", 2, 2 + shift])

    def invoke(self, *arguments):
        return subprocess.run(
            [sys.executable, str(PROJECT_ROOT / "app.py"), "compare-predictions",
             str(self.baseline), str(self.candidate), *map(str, arguments)],
            cwd=self.root, capture_output=True, text=True, encoding="utf-8", timeout=30,
        )

    def test_success_writes_matching_stdout_and_report_without_changing_inputs(self):
        original = (self.baseline.read_bytes(), self.candidate.read_bytes())
        process = self.invoke("--report", self.report)
        self.assertEqual(process.returncode, 0, process.stderr)
        stdout = json.loads(process.stdout)
        self.assertEqual(stdout, json.loads(self.report.read_text(encoding="utf-8")))
        self.assertEqual(stdout["status"], "passed")
        self.assertFalse(stdout["production_accuracy_verified"])
        self.assertEqual(original, (self.baseline.read_bytes(), self.candidate.read_bytes()))

    def test_failed_parity_exits_one_and_preserves_failure_evidence(self):
        self.write_predictions(self.candidate, shift=1)
        process = self.invoke("--report", self.report)
        self.assertEqual(process.returncode, 1, process.stderr)
        report = json.loads(self.report.read_text(encoding="utf-8"))
        self.assertEqual(report, json.loads(process.stdout))
        self.assertEqual(report["status"], "failed")
        self.assertEqual(report["overall"]["prediction_difference"]["rows_outside_tolerance"], 2)

    def test_invalid_artifact_exits_nonzero_without_publishing_success(self):
        self.candidate.write_text("wrong,columns\n1,2\n", encoding="utf-8")
        process = self.invoke("--report", self.report)
        self.assertNotEqual(process.returncode, 0)
        self.assertIn("missing required columns", process.stderr)
        self.assertFalse(self.report.exists())
        self.assertNotIn('"status": "passed"', process.stdout)

    def test_report_cannot_overwrite_either_input(self):
        for input_path in (self.baseline, self.candidate):
            with self.subTest(input=input_path.name):
                original = input_path.read_bytes()
                process = self.invoke("--report", input_path)
                self.assertNotEqual(process.returncode, 0)
                self.assertIn("must not overwrite", process.stderr)
                self.assertEqual(input_path.read_bytes(), original)

    def test_report_symlink_cannot_alias_an_input(self):
        try:
            self.report.symlink_to(self.baseline)
        except OSError as exc:
            self.skipTest(f"Filesystem does not permit symlinks: {exc}")
        original = self.baseline.read_bytes()
        process = self.invoke("--report", self.report)
        self.assertNotEqual(process.returncode, 0)
        self.assertEqual(self.baseline.read_bytes(), original)

    def test_report_temporary_path_cannot_overwrite_an_input(self):
        self.baseline = self.report.with_name(self.report.name + ".tmp")
        self.write_predictions(self.baseline)
        original = self.baseline.read_bytes()
        process = self.invoke("--report", self.report)
        self.assertNotEqual(process.returncode, 0)
        self.assertTrue(self.baseline.exists())
        self.assertEqual(self.baseline.read_bytes(), original)

    def test_invalid_tolerance_exits_nonzero(self):
        process = self.invoke("--atol", "-1", "--report", self.report)
        self.assertNotEqual(process.returncode, 0)
        self.assertIn("finite nonnegative", process.stderr)
        self.assertFalse(self.report.exists())

    def test_cli_help_does_not_import_optional_training_dependencies(self):
        script = textwrap.dedent("""\
            import importlib.abc
            import sys

            blocked = {'numpy', 'pandas', 'torch', 'sklearn', 'xgboost', 'optuna', 'selenium'}

            class BlockTrainingDependencies(importlib.abc.MetaPathFinder):
                def find_spec(self, fullname, path=None, target=None):
                    if fullname.split('.')[0] in blocked:
                        raise AssertionError('CLI help imported optional dependency: ' + fullname)
                    return None

            sys.meta_path.insert(0, BlockTrainingDependencies())
            sys.path.insert(0, sys.argv[1])
            from solar_forecast.cli import main
            main(['compare-predictions', '--help'])
            """)
        process = subprocess.run(
            [sys.executable, "-c", script, str(PROJECT_ROOT / "src")],
            cwd=self.root, capture_output=True, text=True, encoding="utf-8", timeout=30,
        )
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertIn("compare-predictions", process.stdout)
        self.assertIn("--report", process.stdout)


if __name__ == "__main__":
    unittest.main()
