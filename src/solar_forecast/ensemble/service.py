from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from .dynamic_gate import ExplainableDynamicGate, normalize_prediction_columns
from .metrics import aggregate_metrics
from solar_forecast.infrastructure.error_report import write_error_report


class HybridExperiment:
    """Coordinates prediction I/O, dynamic gating, metrics, and artifacts."""

    def __init__(self, output_dir: Path, artifact_level: str = "minimal"):
        if artifact_level not in {"minimal", "standard", "debug"}:
            raise ValueError("artifact_level must be one of: minimal, standard, debug")
        self.output_dir = output_dir
        self.artifact_level = artifact_level

    def run(self, validation_path: Path, test_path: Path) -> dict[str, Path]:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        stage = "load_predictions"
        try:
            validation = normalize_prediction_columns(pd.read_csv(validation_path))
            test = normalize_prediction_columns(pd.read_csv(test_path))
            stage = "fit_dynamic_gate"
            gate = ExplainableDynamicGate().fit(validation)
            stage = "hybrid_evaluation"
            predictions = gate.predict(test)
            metrics = aggregate_metrics(predictions.rename(columns={"hybrid_pred": "y_pred"}))
            assert gate.profiles_ is not None
            return self._write_artifacts(gate.profiles_, predictions, metrics)
        except Exception as exc:
            write_error_report(
                self.output_dir, exc, stage=stage,
                context={"validation_path": validation_path, "test_path": test_path},
            )
            raise

    def _write_artifacts(self, profiles: pd.DataFrame, predictions: pd.DataFrame, metrics: dict[str, pd.DataFrame]) -> dict[str, Path]:
        paths = {
            "gate": self.output_dir / "dynamic_gate_profiles.csv",
            "plant_metrics": self.output_dir / "plant_metrics.csv",
            "region_metrics": self.output_dir / "region_metrics.csv",
            "national_metrics": self.output_dir / "national_metrics.csv",
        }
        profiles.to_csv(paths["gate"], index=False, encoding="utf-8-sig")
        if self.artifact_level in {"standard", "debug"}:
            paths["predictions"] = self.output_dir / "hybrid_predictions.csv.gz"
            predictions.to_csv(paths["predictions"], index=False, compression="gzip")
        for level in ("plant", "region", "national"):
            metrics[level].to_csv(paths[f"{level}_metrics"], index=False, encoding="utf-8-sig")
        manifest = {"status": "completed", "artifact_level": self.artifact_level, "files": {k: str(v) for k, v in paths.items()}}
        (self.output_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        return paths


def build_hybrid_experiment(validation_path: Path, test_path: Path, output_dir: Path, artifact_level: str = "minimal") -> dict[str, Path]:
    """Compatibility facade for callers that have not migrated to HybridExperiment."""
    return HybridExperiment(output_dir, artifact_level).run(validation_path, test_path)
