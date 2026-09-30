"""Check module ownership, naming, imports, generated assets, and source/artifact separation."""
from __future__ import annotations

import ast
import json
from pathlib import Path
import re

from build_dashboard_assets import build_dashboard_bundle
from update_code_index import render_code_index

PROJECT_ROOT = Path(__file__).resolve().parents[1]
GENERIC_MODULE_NAMES = {"service", "config", "models", "model", "data", "base", "utils", "helpers", "common"}
SOURCE_ARTIFACT_SUFFIXES = {".csv", ".gz", ".parquet", ".pt", ".pth", ".pkl", ".pickle", ".joblib", ".sqlite3", ".db", ".jsonl", ".html"}


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
        if path.is_dir() and path.name not in model_ids | {"shared", "__pycache__"}:
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


def main() -> None:
    errors = check_structure()
    if errors:
        raise SystemExit("\n".join(errors))
    print("Structure verified: module purposes, naming, imports, artifact boundaries, and generated files")


if __name__ == "__main__":
    main()
