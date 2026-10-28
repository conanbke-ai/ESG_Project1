"""로컬 실행 단계·실패 중단 계약 검증; 모델 정확도는 모의 검증하지 않는다."""
from __future__ import annotations

import argparse
from contextlib import redirect_stderr, redirect_stdout
import importlib.util
import io
import json
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch


TOOL = Path(__file__).resolve().parents[1] / "tools/run_local_benchmark.py"
spec = importlib.util.spec_from_file_location("local_benchmark_runner", TOOL)
runner = importlib.util.module_from_spec(spec)
spec.loader.exec_module(runner)


class LocalBenchmarkRunnerTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.data = self.root / "gold/model_ready_parts"
        self.data.mkdir(parents=True)
        (self.data / "part.csv").write_text("fixture\n1\n")
        self.manifest = self.data.parent / "model_ready_manifest.json"
        self.manifest.write_text(json.dumps({"contract": runner.GOLD_CONTRACT, "schema_version": 1,
                                            "preprocessing_contract": runner.PREPROCESSING_CONTRACT}))
        self.config = self.root / "config/experiments/calendar.json"
        self.config.parent.mkdir(parents=True)
        self.values = {"input_dataset": "old/location", "output_root": "artifacts/benchmarks",
                       "split": {"train_end": "2024-06-30T23:00:00"},
                       "models": {"cnn_bilstm": {"max_trials_per_candidate": 5}}}
        self.config.write_text(json.dumps(self.values))
        self.run_dir = self.root / "artifacts/benchmarks/exact_returned_run"
        self.args = argparse.Namespace(config=self.config, data=self.data, threads=4,
                                       preflight_only=False, replay_run=None)
        self.calls = []
        self.replay_passes = True
        self.readiness_passes = True
        self.console = io.StringIO()
        self.addCleanup(patch.stopall)
        self.dependencies = patch.object(runner, "_load_dependencies", return_value={"fixture_runtime": True}).start()
        patch.object(runner, "_configure_environment").start()
        patch.object(sys, "path", list(sys.path)).start()
        owner = self

        def readiness(config, **kwargs):
            owner.calls.append("readiness")
            owner.resolved_values = json.loads(config.read_text())
            return {"coverage_gate_passed": owner.readiness_passes}

        class TrainingFixture:
            def __init__(self, **kwargs):
                pass

            def run(self, config, *, smoke):
                owner.assertFalse(smoke)
                owner.calls.append("training")
                owner.run_dir.mkdir(parents=True, exist_ok=True)
                return owner.run_dir

        def replay(run_dir, data):
            owner.assertEqual(run_dir, owner.run_dir)
            owner.assertEqual(data, owner.data)
            owner.calls.append("replay")
            return {"status": "passed" if owner.replay_passes else "failed", "fixture_only": True}

        class DashboardFixture:
            def __init__(self, project_root):
                owner.assertEqual(project_root, owner.root)

            def build(self):
                owner.calls.append("dashboard")
                path = owner.root / "dashboard/data/dashboard_data.json"
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(json.dumps({"model_benchmark": {"run_id": owner.run_dir.name}}))
                return SimpleNamespace(data_path=path)

        patch.dict(sys.modules, {
            "solar_forecast.evaluation.forecast_readiness": SimpleNamespace(run_forecast_readiness=readiness),
            "solar_forecast.jobs.benchmark_job": SimpleNamespace(BenchmarkService=TrainingFixture),
            "verify_benchmark_model_artifacts": SimpleNamespace(verify_selected_artifacts=replay),
            "solar_forecast.reporting.dashboard_builder": SimpleNamespace(DashboardBuilder=DashboardFixture),
        }).start()

    def _execute(self):
        with redirect_stdout(self.console), redirect_stderr(self.console):
            code, report_path = runner.run_local_benchmark(self.args, project_root=self.root)
        return code, json.loads(report_path.read_text())

    def test_full_run_preserves_budget_and_uses_exact_returned_run(self):
        before = self.config.read_bytes()
        code, report = self._execute()
        self.assertEqual(code, 0)
        self.assertEqual(self.calls, ["readiness", "training", "replay", "dashboard"])
        expected = {**self.values, "input_dataset": str(self.data)}
        self.assertEqual(self.resolved_values, expected)
        self.assertEqual(self.config.read_bytes(), before)
        self.assertEqual(report["benchmark_run_dir"], str(self.run_dir))
        self.assertEqual(report["status"], "completed")

    def test_preflight_never_calls_training_replay_or_dashboard(self):
        self.args.preflight_only = True
        code, report = self._execute()
        self.assertEqual(code, 0)
        self.assertEqual(self.calls, ["readiness"])
        self.assertFalse(report["prediction_or_training_performed"])
        self.assertFalse(report["training_invoked"])
        self.assertFalse(report["replay_invoked"])
        self.dependencies.assert_called_once_with(4, require_cuda=True)

    def test_missing_dependency_fails_before_readiness_or_training(self):
        self.dependencies.side_effect = RuntimeError("torch dependency unavailable")
        code, report = self._execute()
        self.assertEqual(code, 1)
        self.assertEqual(self.calls, [])
        self.assertEqual(report["stage"], "dependencies")
        self.assertIn("torch", report["error"]["message"])

    def test_old_gold_fails_before_importing_frameworks(self):
        values = json.loads(self.manifest.read_text())
        values["preprocessing_contract"] = "old"
        self.manifest.write_text(json.dumps(values))
        code, report = self._execute()
        self.assertEqual(code, 1)
        self.dependencies.assert_not_called()
        self.assertEqual(report["stage"], "input_validation")
        self.assertEqual(self.calls, [])

    def test_bad_coverage_never_starts_training(self):
        self.readiness_passes = False
        code, report = self._execute()
        self.assertEqual(code, 1)
        self.assertEqual(self.calls, ["readiness"])
        self.assertFalse(report["training_invoked"])

    def test_failed_replay_keeps_completed_run_and_skips_dashboard(self):
        self.replay_passes = False
        code, report = self._execute()
        self.assertEqual(code, 1)
        self.assertEqual(self.calls, ["readiness", "training", "replay"])
        self.assertEqual(report["benchmark_run_dir"], str(self.run_dir))
        self.assertEqual(report["stage"], "saved_model_replay")
        self.assertEqual(json.loads(Path(report["replay_report"]).read_text())["status"], "failed")

    def test_replay_mode_uses_saved_experiment_without_audit_or_training(self):
        self.run_dir.mkdir(parents=True)
        saved_values = {**self.values, "input_dataset": str(self.data)}
        (self.run_dir / "experiment.json").write_text(json.dumps(saved_values))
        self.config.unlink()
        self.args.replay_run = self.run_dir
        self.args.data = None
        code, report = self._execute()
        self.assertEqual(code, 0)
        self.assertEqual(self.calls, ["replay", "dashboard"])
        self.assertFalse(report["training_invoked"])
        self.assertEqual(report["source_config"], str(self.run_dir / "experiment.json"))
        self.dependencies.assert_called_once_with(4, require_cuda=False)

    def test_unsupported_dashboard_output_fails_before_training(self):
        self.values["output_root"] = "elsewhere"
        self.config.write_text(json.dumps(self.values))
        code, report = self._execute()
        self.assertEqual(code, 1)
        self.assertEqual(self.calls, [])
        self.dependencies.assert_not_called()
        self.assertEqual(report["stage"], "input_validation")


class LocalGpuRequirementTests(unittest.TestCase):
    def test_missing_cuda_rejects_training_before_thread_setup(self):
        torch = MagicMock(__version__="fixture")
        torch.cuda.is_available.return_value = False
        with patch.object(runner.importlib, "import_module", return_value=torch):
            with self.assertRaisesRegex(RuntimeError, "사용자 로컬 GPU"):
                runner._load_dependencies(4)
        torch.set_num_threads.assert_not_called()

    def test_cpu_remains_available_for_replay_without_training(self):
        torch = MagicMock(__version__="fixture")
        torch.cuda.is_available.return_value = False
        torch.get_num_threads.return_value = 4
        torch.get_num_interop_threads.return_value = 1
        with patch.object(runner.importlib, "import_module", return_value=torch):
            result = runner._load_dependencies(4, require_cuda=False)
        self.assertEqual(result["training_device_policy"], "training_not_requested")
        self.assertFalse(result["cuda_available"])
        self.assertEqual(result["replay_device"], "cpu")


if __name__ == "__main__":
    unittest.main()
