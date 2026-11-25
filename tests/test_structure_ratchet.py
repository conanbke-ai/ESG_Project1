"""Exercise the exact diagnostic ratchet independently of existing debt."""
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


@pytest.fixture
def checkout(tmp_path):
    source = tmp_path / "src/solar_forecast/models"
    source.mkdir(parents=True)
    architecture = tmp_path / "config/architecture"
    architecture.mkdir(parents=True)
    (architecture / "modules.json").write_text('{"modules":{}}', encoding="utf-8")
    existing = source.parent / "unregistered.py"
    existing.write_text('"""Existing implementation."""\n', encoding="utf-8")
    baseline = tmp_path / "baseline.json"
    strict = _run(tmp_path)
    assert strict.returncode == 1
    assert "Module lacks an explicit purpose: unregistered.py" in strict.stderr
    _baseline(baseline, strict.stderr.splitlines())
    return tmp_path, baseline, existing, strict.stderr.splitlines()


def test_strict_still_fails_and_unchanged_baseline_is_accepted(checkout):
    root, baseline, _, _ = checkout
    ratchet = _run(root, "--baseline", baseline)
    assert ratchet.returncode == 0, ratchet.stderr
    assert "0 new" in ratchet.stdout
    assert _run(root).returncode == 1


def test_same_total_cannot_hide_new_violation(checkout):
    root, baseline, original, _ = checkout
    original.rename(original.with_name("replacement.py"))
    result = _run(root, "--baseline", baseline)
    assert result.returncode == 1
    assert "New finding: Module lacks an explicit purpose: replacement.py" in result.stderr
    assert "Remove resolved baseline entry: Module lacks an explicit purpose: unregistered.py" in result.stderr


def test_resolved_debt_must_be_removed_and_cannot_return(checkout):
    root, baseline, original, findings = checkout
    catalog = root / "config/architecture/modules.json"
    catalog.write_text('{"modules":{"unregistered.py":{"purpose":"example"}}}', encoding="utf-8")
    assert "Remove resolved baseline entry" in _run(root, "--baseline", baseline).stderr
    _baseline(baseline, [item for item in findings if "unregistered.py" not in item])
    assert _run(root, "--baseline", baseline).returncode == 0
    catalog.write_text('{"modules":{}}', encoding="utf-8")
    assert "New finding: Module lacks an explicit purpose: unregistered.py" in _run(root, "--baseline", baseline).stderr


def test_wrong_baseline_contract_fails_closed(checkout):
    root, baseline, _, findings = checkout
    _baseline(baseline, findings, "other-version")
    result = _run(root, "--baseline", baseline)
    assert result.returncode == 1
    assert "Invalid structure baseline" in result.stderr
