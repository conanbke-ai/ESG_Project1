"""Check module ownership, naming, imports, generated assets, and source/artifact separation."""
from __future__ import annotations

import argparse
import ast
from collections import Counter
from hashlib import sha256
import json
from pathlib import Path
import re
import subprocess

from build_dashboard_assets import build_dashboard_bundle
from update_code_index import render_code_index

PROJECT_ROOT = Path(__file__).resolve().parents[1]
BASELINE_CONTRACT = "solar-structure-baseline.v2"
RULES_VERSION = "solar-structure-rules.v1"
GENERIC_MODULE_NAMES = {"service", "config", "models", "model", "data", "base", "utils", "helpers", "common"}
SOURCE_ARTIFACT_SUFFIXES = {".csv", ".gz", ".parquet", ".pt", ".pth", ".pkl", ".pickle", ".joblib", ".sqlite3", ".db", ".jsonl", ".html"}
PYTHON_BYTECODE_SUFFIXES = {".pyc", ".pyo"}


def is_generated_python_cache_tree(path: Path) -> bool:
    """Return whether a directory contains only Python bytecode cache entries."""
    if not path.is_dir():
        return False
    entries = list(path.rglob("*"))
    if not entries:
        return False
    return all(
        (entry.is_dir() and entry.name == "__pycache__")
        or (entry.is_file() and entry.suffix.lower() in PYTHON_BYTECODE_SUFFIXES)
        for entry in entries
    )


def module_exports(tree: ast.Module) -> set[str]:
    """Resolve declared Python exports without importing optional ML frameworks."""
    names = set()
    for node in tree.body:
        if isinstance(node, (ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
            names.add(node.name)
        elif isinstance(node, (ast.Import, ast.ImportFrom)):
            names.update(alias.asname or alias.name.split(".")[0] for alias in node.names)
        elif isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            names.update(target.id for target in targets if isinstance(target, ast.Name))
            if any(isinstance(target, ast.Name) and target.id in {"_EXPORTS", "__all__"} for target in targets):
                try:
                    names.update(ast.literal_eval(node.value))
                except (ValueError, TypeError):
                    pass
    return names


def check_structure(project_root: Path = PROJECT_ROOT) -> list[str]:
    source_root = project_root / "src/solar_forecast"
    catalog = json.loads((project_root / "config/architecture/modules.json").read_text(encoding="utf-8"))["modules"]
    files = {path.relative_to(source_root).as_posix(): path for path in source_root.rglob("*.py")}
    errors = []
    # Compatibility modules are temporary import bridges, never canonical feature owners.
    for relative, entry in sorted(catalog.items()):
        if entry.get("kind") != "compatibility_shim":
            continue
        source = files.get(relative)
        if source is None:
            continue  # Missing-file validation below owns this error.
        old_module = "solar_forecast." + relative.removesuffix(".py").replace("/", ".")
        move_catalog = json.loads(
            (project_root / "config/architecture/module_moves.json").read_text(encoding="utf-8")
        )["modules"]
        target_relative = move_catalog.get(relative)
        if not target_relative or target_relative == relative or target_relative not in files:
            errors.append(f"Compatibility shim lacks a canonical move target: {relative}")
            continue
        canonical_module = "solar_forecast." + target_relative.removesuffix(".py").replace("/", ".")
        if canonical_module not in source.read_text(encoding="utf-8"):
            errors.append(f"Compatibility shim does not import canonical owner: {relative} -> {canonical_module}")
        if old_module == canonical_module:
            errors.append(f"Compatibility shim points to itself: {relative}")
    for missing in sorted(files.keys() - catalog.keys()):
        errors.append(f"Module lacks an explicit purpose: {missing}")
    for stale in sorted(catalog.keys() - files.keys()):
        errors.append(f"Catalog references a missing module: {stale}")
    trees = {relative: ast.parse(path.read_text(encoding="utf-8"), filename=str(path)) for relative, path in files.items()}
    exports = {relative: module_exports(tree) for relative, tree in trees.items()}
    for relative, tree in trees.items():
        path = files[relative]
        if not re.fullmatch(r"[a-z][a-z0-9_]*|__[a-z]+__", path.stem) or path.stem in GENERIC_MODULE_NAMES:
            errors.append(f"Use a role-specific snake_case module name: {relative}")
        if not ast.get_docstring(tree):
            errors.append(f"Module lacks a responsibility docstring: {relative}")
        if path.parent == source_root / "models" and path.name != "__init__.py":
            errors.append(f"Place model implementation under models/<model_id>/: {relative}")
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and not re.fullmatch(r"_*[a-z][a-z0-9_]*_*", node.name):
                errors.append(f"Use snake_case function names: {relative}:{node.name}")
            if not isinstance(node, ast.ImportFrom) or not node.module or not node.module.startswith("solar_forecast"):
                continue
            module_path = node.module.removeprefix("solar_forecast").lstrip(".").replace(".", "/")
            candidates = [module_path + ".py", (module_path + "/" if module_path else "") + "__init__.py"]
            target = next((candidate for candidate in candidates if candidate in trees), None)
            if target is None:
                errors.append(f"Unresolved import: {relative} -> {node.module}")
                continue
            for alias in node.names:
                child = module_path + "/" + alias.name
                if alias.name != "*" and alias.name not in exports[target] and child + ".py" not in trees and child + "/__init__.py" not in trees:
                    errors.append(f"Unresolved imported symbol: {relative} -> {node.module}:{alias.name}")
            if not relative.startswith("cli/") and node.module.startswith("solar_forecast.cli"):
                errors.append(f"Core code must not depend on CLI handlers: {relative}")
    for path in source_root.rglob("*"):
        if path.is_file() and path.suffix.lower() in SOURCE_ARTIFACT_SUFFIXES:
            errors.append(f"Runtime artifact inside source tree: {path.relative_to(project_root)}")
        if path.is_file() and path.suffix == ".json" and path.relative_to(source_root).as_posix() != "collectors/sources.json":
            errors.append(f"JSON settings/artifacts belong outside source, except the provider catalog: {path.relative_to(project_root)}")
    model_ids = {path.stem for path in (project_root / "config/models").glob("*.json")}
    for path in (source_root / "models").iterdir():
        if (
            path.is_dir()
            and path.name not in model_ids | {"shared", "__pycache__"}
            and not is_generated_python_cache_tree(path)
        ):
            errors.append(f"Model folder needs a matching config/models/<model_id>.json: {path.name}")
    for path in project_root.glob("*.py"):
        if path.name != "app.py":
            errors.append(f"Root Python scripts belong in tools/ or src/: {path.name}")
    dashboard = project_root / "dashboard/assets/dashboard.js"
    if not dashboard.exists() or dashboard.read_text(encoding="utf-8") != build_dashboard_bundle(project_root):
        errors.append("Dashboard bundle is stale: run python tools/build_dashboard_assets.py")
    index = project_root / "docs/CODE_INDEX.md"
    if not index.exists() or index.read_text(encoding="utf-8") != render_code_index(project_root):
        errors.append("Code index is stale: run python tools/update_code_index.py")
    return errors


def read_approved_baseline(baseline_path: Path, expected_sha256: str) -> list[str]:
    """Use the verifier's separately approved digest, never trust a self-declared hash."""
    if not re.fullmatch(r"[0-9a-fA-F]{64}", expected_sha256):
        raise ValueError("Approved baseline SHA-256 must have 64 hexadecimal digits")
    payload = baseline_path.read_bytes()
    if sha256(payload).hexdigest() != expected_sha256.lower():
        raise ValueError("Baseline differs from the approved SHA-256")
    baseline = json.loads(payload.decode("utf-8"))
    if not isinstance(baseline, dict) or baseline.get("contract") not in {
        "solar-structure-baseline.v1", BASELINE_CONTRACT,
    }:
        raise ValueError("Unsupported structure baseline contract")
    if baseline["contract"] == BASELINE_CONTRACT:
        if baseline.get("rules_version") != RULES_VERSION:
            raise ValueError("Baseline rules version does not match this checker")
        source_sha = baseline.get("source_sha")
        if (not isinstance(source_sha, str) or not re.fullmatch(r"(?:[0-9a-f]{40}|[0-9a-f]{64})", source_sha)
                or type(baseline.get("source_dirty")) is not bool):
            raise ValueError("Baseline source provenance is invalid")
    recorded = baseline.get("findings")
    if not isinstance(recorded, list) or any(
        not isinstance(item, str) or not item.strip() or "\n" in item or "\r" in item
        for item in recorded
    ):
        raise ValueError("Structure baseline findings must be nonempty single-line strings")
    return recorded


def compare_structure_baseline(findings: list[str], approved: list[str]) -> tuple[list[str], list[str]]:
    """Compare full diagnostic identities and multiplicities without masking findings."""
    current, previous = Counter(findings), Counter(approved)
    return sorted((current - previous).elements()), sorted((previous - current).elements())


def capture_structure_baseline(
    root: Path, destination: Path, findings: list[str], *, predecessor_sha256: str | None = None,
) -> str:
    """Write a new review artifact, separate from the approved input baseline."""
    def git(*args: str) -> str:
        return subprocess.run(
            ["git", "--no-optional-locks", "-C", str(root), *args],
            check=True, capture_output=True, text=True, encoding="utf-8",
        ).stdout.strip()

    if Path(git("rev-parse", "--show-toplevel")).resolve() != root:
        raise ValueError("Capture requires the root of the Git checkout being inspected")
    document = {
        "contract": BASELINE_CONTRACT,
        "rules_version": RULES_VERSION,
        "source_sha": git("rev-parse", "HEAD"),
        "source_dirty": bool(git("status", "--porcelain", "--untracked-files=normal")),
        "findings": sorted(findings),
    }
    if predecessor_sha256 is not None:
        document["predecessor_sha256"] = predecessor_sha256.lower()
    payload = (json.dumps(document, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
    with destination.open("xb") as stream:
        stream.write(payload)
    return sha256(payload).hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", type=Path, default=PROJECT_ROOT,
                        help="Project checkout to inspect (defaults to this checkout)")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--baseline", type=Path, help="Compare with a separately approved baseline")
    mode.add_argument("--capture-baseline", type=Path, help="Capture a new snapshot for review, not a gate result")
    parser.add_argument("--baseline-sha256", help="Approved SHA-256 obtained outside the candidate changes")
    parser.add_argument("--record-reduction", type=Path, help="Write a new baseline only when verified debt decreases")
    args = parser.parse_args()
    if args.baseline is not None and not args.baseline_sha256:
        parser.error("--baseline requires a separately approved --baseline-sha256")
    if args.baseline is None and (args.baseline_sha256 is not None or args.record_reduction is not None):
        parser.error("--baseline-sha256 and --record-reduction require --baseline")
    root = args.project_root.resolve()
    errors = check_structure(root)
    if args.capture_baseline is not None:
        try:
            digest = capture_structure_baseline(root, args.capture_baseline, errors)
        except (ValueError, OSError, subprocess.CalledProcessError) as exc:
            raise SystemExit(f"Cannot capture structure baseline: {exc}") from exc
        print(f"Structure snapshot for review: {len(errors)} findings, SHA-256 {digest}")
        return
    if args.baseline is not None:
        try:
            approved = read_approved_baseline(args.baseline, args.baseline_sha256)
        except (ValueError, OSError) as exc:
            raise SystemExit(f"Invalid structure baseline: {exc}") from exc
        new, retired = compare_structure_baseline(errors, approved)
        if new or (retired and args.record_reduction is None):
            lines = [*(f"New finding: {item}" for item in new),
                     *(f"Remove resolved baseline entry: {item}" for item in retired)]
            raise SystemExit("\n".join(lines))
        if args.record_reduction is not None:
            if not retired:
                raise SystemExit("Structure debt did not decrease")
            try:
                digest = capture_structure_baseline(
                    root, args.record_reduction, errors, predecessor_sha256=args.baseline_sha256,
                )
            except (ValueError, OSError, subprocess.CalledProcessError) as exc:
                raise SystemExit(f"Cannot record structure debt reduction: {exc}") from exc
            print(f"Structure debt decreased by {len(retired)} to {len(errors)} findings, SHA-256 {digest}")
            return
        print(f"Structure baseline matched: {len(errors)} existing findings, 0 new")
        return
    if errors:
        raise SystemExit("\n".join(errors))
    print("Structure verified: module purposes, naming, imports, artifact boundaries, and generated files")


if __name__ == "__main__":
    main()
