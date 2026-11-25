"""Exercise the exact diagnostic ratchet independently of existing debt."""
from hashlib import sha256
import json
from pathlib import Path
import subprocess
import sys

import pytest


ROOT = Path(__file__).resolve().parents[1]


def _run(root, *args):
    return subprocess.run(
        [sys.executable, "-X", "utf8", "-B", str(ROOT / "tools/check_structure.py"),
         "--project-root", str(root), *map(str, args)],
        text=True, capture_output=True, check=False, encoding="utf-8",
    )


def _baseline(path, findings, contract="solar-structure-baseline.v1"):
    path.write_text(json.dumps({"contract": contract, "findings": findings}), encoding="utf-8")
    return sha256(path.read_bytes()).hexdigest()


def _git(root, *args):
    return subprocess.check_output(["git", "-C", str(root), *args], text=True).strip()


@pytest.fixture
def checkout(tmp_path):
    source = tmp_path / "src/solar_forecast/models"
    source.mkdir(parents=True)
    architecture = tmp_path / "config/architecture"
    architecture.mkdir(parents=True)
    (architecture / "modules.json").write_text('{"modules":{}}', encoding="utf-8")
    existing = source.parent / "unregistered.py"
    existing.write_text('"""Existing implementation."""\n', encoding="utf-8")
    _git(tmp_path, "init", "-q")
    _git(tmp_path, "add", "src", "config")
    _git(tmp_path, "-c", "user.name=Fixture", "-c", "user.email=fixture@example.invalid",
         "commit", "-qm", "Test source")
    baseline = tmp_path / "baseline.json"
    strict = _run(tmp_path)
    assert strict.returncode == 1
    assert "Module lacks an explicit purpose: unregistered.py" in strict.stderr
    _baseline(baseline, strict.stderr.splitlines())
    return tmp_path, baseline, existing, strict.stderr.splitlines()


def test_strict_still_fails_and_unchanged_baseline_is_accepted(checkout):
    root, baseline, _, _ = checkout
    unpinned = _run(root, "--baseline", baseline)
    assert unpinned.returncode != 0
    assert "separately approved" in unpinned.stderr
    ratchet = _run(root, "--baseline", baseline, "--baseline-sha256", sha256(baseline.read_bytes()).hexdigest())
    assert ratchet.returncode == 0, ratchet.stderr
    assert "0 new" in ratchet.stdout
    assert _run(root).returncode == 1


def test_same_total_cannot_hide_new_violation(checkout):
    root, baseline, original, _ = checkout
    digest = sha256(baseline.read_bytes()).hexdigest()
    original.rename(original.with_name("replacement.py"))
    result = _run(root, "--baseline", baseline, "--baseline-sha256", digest)
    assert result.returncode == 1
    assert "New finding: Module lacks an explicit purpose: replacement.py" in result.stderr
    assert "Remove resolved baseline entry: Module lacks an explicit purpose: unregistered.py" in result.stderr
    reduced = root / "reduced.json"
    blocked = _run(root, "--baseline", baseline, "--baseline-sha256", digest, "--record-reduction", reduced)
    assert blocked.returncode == 1 and not reduced.exists()


def test_resolved_debt_must_be_removed_and_cannot_return(checkout):
    root, baseline, _, _ = checkout
    approved_bytes = baseline.read_bytes()
    digest = sha256(approved_bytes).hexdigest()
    catalog = root / "config/architecture/modules.json"
    catalog.write_text('{"modules":{"unregistered.py":{"purpose":"example"}}}', encoding="utf-8")
    args = ("--baseline", baseline, "--baseline-sha256", digest)
    assert "Remove resolved baseline entry" in _run(root, *args).stderr
    reduced = root / "reduced.json"
    result = _run(root, *args, "--record-reduction", reduced)
    assert result.returncode == 0, result.stderr
    assert baseline.read_bytes() == approved_bytes
    evidence = json.loads(reduced.read_text(encoding="utf-8"))
    assert evidence["predecessor_sha256"] == digest
    assert all("unregistered.py" not in item for item in evidence["findings"])
    reduced_args = ("--baseline", reduced, "--baseline-sha256", sha256(reduced.read_bytes()).hexdigest())
    assert _run(root, *reduced_args).returncode == 0
    catalog.write_text('{"modules":{}}', encoding="utf-8")
    assert "New finding: Module lacks an explicit purpose: unregistered.py" in _run(root, *reduced_args).stderr


def test_finding_added_to_baseline_is_rejected_by_independent_digest(checkout):
    root, baseline, _, findings = checkout
    approved_digest = sha256(baseline.read_bytes()).hexdigest()
    (root / "src/solar_forecast/new_module.py").write_text('"""New module."""\n', encoding="utf-8")
    _baseline(baseline, [*findings, "Module lacks an explicit purpose: new_module.py"])
    result = _run(root, "--baseline", baseline, "--baseline-sha256", approved_digest)
    assert result.returncode == 1
    assert "Baseline differs from the approved SHA-256" in result.stderr


def test_capture_records_git_source_and_rules_without_overwriting(checkout):
    root, baseline, _, findings = checkout
    before = baseline.read_bytes()
    assert _run(root, "--capture-baseline", baseline).returncode == 1
    assert baseline.read_bytes() == before
    snapshot = root / "snapshot.json"
    result = _run(root, "--capture-baseline", snapshot)
    assert result.returncode == 0, result.stderr
    document = json.loads(snapshot.read_text(encoding="utf-8"))
    assert document["contract"] == "solar-structure-baseline.v2"
    assert document["rules_version"] == "solar-structure-rules.v1"
    assert document["source_sha"] == _git(root, "rev-parse", "HEAD")
    assert document["source_dirty"] is True
    assert document["findings"] == sorted(findings)
    assert sha256(snapshot.read_bytes()).hexdigest() in result.stdout


def test_snapshot_from_different_rules_is_rejected(checkout):
    root, baseline, _, _ = checkout
    snapshot = root / "snapshot.json"
    assert _run(root, "--capture-baseline", snapshot).returncode == 0
    document = json.loads(snapshot.read_text(encoding="utf-8"))
    document["rules_version"] = "solar-structure-rules.v999"
    snapshot.write_text(json.dumps(document), encoding="utf-8")
    result = _run(root, "--baseline", snapshot, "--baseline-sha256", sha256(snapshot.read_bytes()).hexdigest())
    assert result.returncode == 1
    assert "Baseline rules version does not match this checker" in result.stderr


def test_wrong_baseline_contract_fails_closed(checkout):
    root, baseline, _, findings = checkout
    digest = _baseline(baseline, findings, "other-version")
    result = _run(root, "--baseline", baseline, "--baseline-sha256", digest)
    assert result.returncode == 1
    assert "Invalid structure baseline" in result.stderr
