#!/usr/bin/env python3
"""TORI cross-project QA evidence contract. Standard library only.

Usage: python tools/qa_contract_gate.py path/to/evidence.json
Exit 0: contract valid; exit 1: contract invalid or non-PASS exit decision.
This validates evidence *structure*, not truth of execution or agent invocation.
"""
import json
import pathlib
import sys

PROJECTS = {"nongtori", "solar", "untori", "mystori", "blotori"}
RISKS = {"LOW", "MEDIUM", "HIGH", "CRITICAL"}
DECISIONS = {"PASS", "PARTIAL", "BLOCKED", "NEEDS_REVISION"}
EVIDENCE_STATES = {"DESIGNED", "REVIEWED", "TEST_EXECUTED", "PASS", "FAIL", "NOT_RUN", "NOT_APPLICABLE"}
REQUIRED = ("feature", "project", "risk", "acceptance_criteria", "qa_roles",
            "test_cases", "regression_scope", "evidence", "not_run", "exit_decision")

def problems(document):
    errors = []
    if not isinstance(document, dict):
        return ["Evidence must be a JSON object"]
    for key in REQUIRED:
        if key not in document:
            errors.append(f"Missing: {key}")
    if errors:
        return errors
    if document["project"] not in PROJECTS:
        errors.append("Unknown project")
    if document["risk"] not in RISKS:
        errors.append("Unknown risk")
    for key in ("feature", "regression_scope"):
        if not isinstance(document[key], str) or not document[key].strip():
            errors.append(f"{key} must be nonempty")
    for key in ("acceptance_criteria", "qa_roles", "test_cases", "evidence", "not_run"):
        if not isinstance(document[key], list):
            errors.append(f"{key} must be a list")
    if errors:
        return errors
    for key in ("acceptance_criteria", "qa_roles", "test_cases"):
        if not document[key] or not all(isinstance(v, str) and v.strip() for v in document[key]):
            errors.append(f"{key} requires nonempty strings")
    if "QA Lead & Test Architect" not in document["qa_roles"]:
        errors.append("QA Lead missing")
    if "Acceptance Reviewer" not in document["qa_roles"]:
        errors.append("Acceptance Reviewer missing")
    if document.get("user_facing") and "Product Flow QA Reviewer" not in document["qa_roles"]:
        errors.append("User-facing change missing Product Flow QA")
    if document["risk"] in {"MEDIUM", "HIGH", "CRITICAL"} and "Functional & Regression QA Reviewer" not in document["qa_roles"]:
        errors.append("Medium+ risk missing Functional & Regression QA")
    executed = False
    for index, record in enumerate(document["evidence"]):
        if not isinstance(record, dict) or record.get("state") not in EVIDENCE_STATES:
            errors.append(f"evidence[{index}]: invalid state")
            continue
        if not isinstance(record.get("check"), str) or not record["check"].strip():
            errors.append(f"evidence[{index}]: check required")
        if record["state"] in {"PASS", "FAIL", "TEST_EXECUTED"}:
            if not isinstance(record.get("proof"), str) or not record["proof"].strip():
                errors.append(f"evidence[{index}]: executable proof required")
            else:
                executed = True
    if document["exit_decision"] not in DECISIONS:
        errors.append("Invalid exit_decision")
    elif document["exit_decision"] != "PASS":
        errors.append(f"Exit decision not PASS: {document['exit_decision']}")
    if document["not_run"]:
        errors.append("Unresolved NOT_RUN; cannot PASS")
    if any(isinstance(r, dict) and r.get("state") == "FAIL" for r in document["evidence"]):
        errors.append("Unresolved FAIL; cannot PASS")
    if not executed:
        errors.append("No executable evidence; REVIEWED/DESIGNED alone cannot PASS")
    return errors

def main(argv):
    if len(argv) < 2:
        print("Usage: qa_contract_gate.py evidence.json [evidence2.json ...]", file=sys.stderr)
        return 2
    failed = False
    for name in argv[1:]:
        try:
            document = json.loads(pathlib.Path(name).read_text(encoding="utf-8"))
            errs = problems(document)
        except (OSError, ValueError) as exc:
            errs = [f"Cannot read JSON: {exc}"]
        for err in errs:
            print(f"[QA CONTRACT FAIL] {name}: {err}", file=sys.stderr)
        if not errs:
            print(f"[QA CONTRACT PASS] {name} (schema and evidence declaration only)")
        failed = failed or bool(errs)
    return 1 if failed else 0

if __name__ == "__main__":
    sys.exit(main(sys.argv))
