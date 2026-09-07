from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

from solar_forecast.artifacts.manifest import write_json_atomic
from solar_forecast.cli import build_parser
from solar_forecast.settings import ModelJobConfig
from solar_forecast.verification import (
    PipelineVerificationService,
    VerificationConfig,
)


def test_verify_e2e_cli_defaults_to_smoke_and_requires_collect_dates() -> None:
    parser = build_parser()

    args = parser.parse_args(["verify-e2e"])
    assert args.full_train is False
    assert args.models == "xgboost,cnn_bilstm"
    assert args.collect is False


def test_verification_service_writes_auditable_report(tmp_path, monkeypatch) -> None:
    collection_dir = tmp_path / "file/raw/runs/20260907_000000_000000"
    collection_dir.mkdir(parents=True)
    write_json_atomic(
        collection_dir / "collection_manifest.json",
        {
            "start_date": "2025-01-01",
            "end_date": "2025-01-02",
            "results": [
                {
                    "source": "koen",
                    "status": "downloaded",
                    "rows": 10,
                    "files": ["koen.csv"],
                    "message": "ok",
                }
            ],
            "file_artifacts": [{"path": "koen.csv"}],
        },
    )
    downloads = tmp_path / "file/standardized/downloads/koen"
    downloads.mkdir(parents=True)
    (downloads / "sample.csv").write_text("a,b\n1,2\n", encoding="utf-8")

    class FakePreparationService:
        def __init__(self, input_root, weather_root, merged_source, output_dir):
            self.output_dir = Path(output_dir)

        def run(self):
            model_path = self.output_dir / "model_ready.csv.gz"
            model_path.parent.mkdir(parents=True, exist_ok=True)
            write_json_atomic(
                model_path.with_name("model_ready_manifest.json"),
                {
                    "training_eligibility": {
                        "training_eligible_rows": 8,
                        "training_eligible_plants": 2,
                    }
                },
            )
            return SimpleNamespace(
                generation=SimpleNamespace(
                    manifest_path=self.output_dir / "generation_manifest.json",
                    partitions=[SimpleNamespace(destination="p1.csv.gz")],
                    rows=10,
                ),
                model_dataset=SimpleNamespace(
                    path=model_path,
                    rows=8,
                    plants=2,
                    partitions_dir=self.output_dir / "model_ready_parts",
                ),
                quality=SimpleNamespace(
                    report_path=self.output_dir / "plant_quality_report.csv",
                    high_risk_plants=0,
                    review_plants=1,
                    preprocessing_artifact_plants=0,
                ),
            )

    class FakeTrainingService:
        def run(self, config, *, smoke):
            run_dir = tmp_path / config.values["output_root"] / "run"
            write_json_atomic(
                run_dir / "manifest.json",
                {
                    "status": "completed",
                    "model": config.model,
                    "run_id": "run",
                    "details": {
                        "metrics": {"mae": 0.1},
                        "run_context": {
                            "execution_mode": "smoke" if smoke else "full"
                        },
                        "memory_aware_loading": {"retained_rows": 8},
                        "checkpoint": {"enabled": True},
                        "test_predictions_path": str(run_dir / "test.csv"),
                    },
                },
            )
            return run_dir

    class FakeDashboardBuilder:
        def __init__(self, project_root, output_dir):
            self.output_dir = Path(output_dir)

        def build(self):
            return SimpleNamespace(
                data_path=self.output_dir / "data/dashboard_data.json",
                boundary_path=self.output_dir / "data/province_boundaries.geojson",
                solar_dashboard=self.output_dir / "solar_dashboard.html",
                forecast_dashboard=self.output_dir / "forecast.html",
                analytics_dashboard=self.output_dir / "model_analysis.html",
                national_generator_records=10,
                national_capacity_mw=1.5,
                model_analysis_status="empty",
                data_quality_signals=1,
            )

    monkeypatch.setattr(
        "solar_forecast.verification.DataPreparationService",
        FakePreparationService,
    )
    monkeypatch.setattr(
        "solar_forecast.verification.TrainingService",
        FakeTrainingService,
    )
    monkeypatch.setattr(
        "solar_forecast.verification.DashboardBuilder",
        FakeDashboardBuilder,
    )
    monkeypatch.setattr(
        "solar_forecast.verification.load_model_config",
        lambda path: ModelJobConfig(
            "xgboost",
            "test",
            {"model": "xgboost", "output_root": "artifacts/models/xgboost"},
            Path(path),
        ),
    )

    result = PipelineVerificationService(
        VerificationConfig(
            project_root=tmp_path,
            train_models=("xgboost",),
            report_root=Path("artifacts/verification/e2e"),
        )
    ).run()

    report = json.loads(result.report_path.read_text(encoding="utf-8"))
    assert result.status == "passed_with_warnings"
    assert report["contract"] == "solar-e2e-verification.v1"
    assert [step["name"] for step in report["steps"]] == [
        "collection_manifest",
        "prepare_data",
        "train_xgboost",
        "build_dashboard",
    ]
    assert report["execution"]["training_mode"] == "smoke"
    assert report["summary"]["warnings"] >= 1
