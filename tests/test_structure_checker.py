"""Structure-policy checks for generated caches and real model source folders."""
from __future__ import annotations

import importlib.util
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch


TOOLS_ROOT = Path(__file__).parents[1] / "tools"
SPEC = importlib.util.spec_from_file_location(
    "structure_checker", TOOLS_ROOT / "check_structure.py"
)
structure_checker = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
with patch.dict(
    "sys.modules",
    {"build_dashboard_assets": Mock(), "update_code_index": Mock()},
):
    SPEC.loader.exec_module(structure_checker)


class StructureCheckerTests(unittest.TestCase):
    def _errors_for_model_folder(self, populate) -> list[str]:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            model_folder = root / "src" / "solar_forecast" / "models" / "cnn"
            model_folder.mkdir(parents=True)
            (root / "config" / "architecture").mkdir(parents=True)
            (root / "config" / "models").mkdir(parents=True)
            (root / "config" / "architecture" / "modules.json").write_text(
                '{"modules": {}}', encoding="utf-8"
            )
            (root / "config" / "architecture" / "module_moves.json").write_text(
                '{"modules": {}}', encoding="utf-8"
            )
            (root / "dashboard" / "assets").mkdir(parents=True)
            (root / "dashboard" / "assets" / "dashboard.js").write_text(
                "generated-dashboard", encoding="utf-8"
            )
            (root / "docs").mkdir()
            (root / "docs" / "CODE_INDEX.md").write_text(
                "generated-index", encoding="utf-8"
            )
            populate(model_folder)
            with (
                patch.object(
                    structure_checker,
                    "build_dashboard_bundle",
                    return_value="generated-dashboard",
                ),
                patch.object(
                    structure_checker,
                    "render_code_index",
                    return_value="generated-index",
                ),
            ):
                return structure_checker.check_structure(root)

    def test_generated_bytecode_only_does_not_create_a_model_folder(self):
        def populate(model_folder: Path) -> None:
            cache = model_folder / "__pycache__"
            cache.mkdir()
            (cache / "legacy_model.cpython-311.pyc").write_bytes(b"bytecode")
            (model_folder / "legacy_model.pyc").write_bytes(b"bytecode")

        self.assertEqual(self._errors_for_model_folder(populate), [])

    def test_legacy_python_source_still_requires_a_model_contract(self):
        def populate(model_folder: Path) -> None:
            cache = model_folder / "__pycache__"
            cache.mkdir()
            (cache / "legacy_model.cpython-311.pyc").write_bytes(b"bytecode")
            (model_folder / "legacy_model.py").write_text(
                '"""Forbidden legacy model owner."""\n', encoding="utf-8"
            )

        errors = self._errors_for_model_folder(populate)
        self.assertIn(
            "Model folder needs a matching config/models/<model_id>.json: cnn",
            errors,
        )
        self.assertIn("Module lacks an explicit purpose: models/cnn/legacy_model.py", errors)

    def test_arbitrary_noncache_file_is_not_ignored(self):
        def populate(model_folder: Path) -> None:
            (model_folder / "notes.txt").write_text("not generated", encoding="utf-8")

        expected = ["Model folder needs a matching config/models/<model_id>.json: cnn"]
        self.assertEqual(self._errors_for_model_folder(lambda _: None), expected)
        self.assertEqual(self._errors_for_model_folder(populate), expected)


if __name__ == "__main__":
    unittest.main()
