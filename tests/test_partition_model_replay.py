"""Dataset identity/replay plumbing tests; no model accuracy is simulated."""
from __future__ import annotations

import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import tempfile
import unittest

from solar_forecast.config_loader import PROJECT_ROOT
from solar_forecast.models.shared.checkpoint_store import dataset_signature, PARTITION_FINGERPRINT_CONTRACT

spec = importlib.util.spec_from_file_location(
    "verify_benchmark_model_artifacts", PROJECT_ROOT / "tools/verify_benchmark_model_artifacts.py"
)
replay = importlib.util.module_from_spec(spec)
spec.loader.exec_module(replay)


class PartitionReplayTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root / "original/model_ready_parts"
        self.source.mkdir(parents=True)
        self.part = self.source / "part.csv"
        self.part.write_text("value\n1\n", encoding="utf-8")
        (self.source.parent / "model_ready_manifest.json").write_text('{"contract":"fixture"}')

    def test_partition_identity_survives_copy_and_timestamp_changes(self):
        expected = dataset_signature(self.source)
        moved = self.root / "moved"
        shutil.copytree(self.source.parent, moved, copy_function=shutil.copyfile)
        os.utime(moved / "model_ready_parts/part.csv", ns=(1, 1))
        self.assertEqual(dataset_signature(moved / "model_ready_parts"), expected)

    def test_same_size_same_mtime_content_change_is_detected(self):
        expected = dataset_signature(self.source)
        previous = self.part.stat()
        self.part.write_text("value\n2\n", encoding="utf-8")
        os.utime(self.part, ns=(previous.st_atime_ns, previous.st_mtime_ns))
        self.assertEqual(self.part.stat().st_size, previous.st_size)
        self.assertNotEqual(dataset_signature(self.source), expected)

    def test_file_fingerprint_remains_plain_sha256(self):
        self.assertEqual(dataset_signature(self.part), hashlib.sha256(self.part.read_bytes()).hexdigest())

    def test_manifest_or_partition_inventory_change_invalidates_identity(self):
        expected = dataset_signature(self.source)
        (self.source / "second.csv").write_text("value\n3\n")
        self.assertNotEqual(dataset_signature(self.source), expected)
        (self.source / "second.csv").unlink()
        (self.source.parent / "model_ready_manifest.json").write_text('{"contract":"changed"}')
        self.assertNotEqual(dataset_signature(self.source), expected)

    def test_missing_and_empty_sources_are_rejected(self):
        with self.assertRaises(FileNotFoundError):
            dataset_signature(self.root / "missing")
        empty = self.root / "empty"
        empty.mkdir()
        with self.assertRaisesRegex(ValueError, "empty"):
            dataset_signature(empty)

    def test_replay_accepts_partition_source_but_no_models_never_passes(self):
        run = self.root / "run"
        run.mkdir()
        manifest = {"status": "completed", "provenance": {
            "dataset_fingerprint": dataset_signature(self.source)}, "tasks": []}
        (run / "manifest.json").write_text(json.dumps(manifest))
        report = replay.verify_selected_artifacts(run, self.source)
        self.assertEqual(report["dataset_fingerprint_contract"], PARTITION_FINGERPRINT_CONTRACT)
        self.assertEqual(report["status"], "failed")
        self.assertFalse(report["training_performed"])
        self.part.write_text("value\n9\n")
        with self.assertRaisesRegex(ValueError, "exact observed dataset"):
            replay.verify_selected_artifacts(run, self.source)


if __name__ == "__main__":
    unittest.main()
