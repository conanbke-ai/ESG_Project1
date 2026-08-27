import numpy as np
import pandas as pd

from solar_forecast.quality import (
    GenerationQualityPolicy,
    PlantQualityProfiler,
    QualityAuditService,
)


def _quality_frame() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "timestamp": pd.date_range("2025-01-01 10:00", periods=6, freq="h"),
            "company": "iwest",
            "plant_id": "iwest:test",
            "plant": "test",
            "region": "A",
            "energy_source": "solar",
            "capacity_mw": 1.0,
            "generation_mwh": [-0.1, 0.0, 0.5, 0.5, 0.5, 0.5],
            "solar_irradiance_mj_m2": [0.5] * 6,
            "humidity_pct": [50, 50, 101, 50, 50, 50],
        }
    )


def test_quality_policy_flags_negative_without_turning_it_into_zero():
    result = GenerationQualityPolicy().apply(_quality_frame())
    assert result.iloc[0]["generation_mwh"] == -0.1
    assert bool(result.iloc[0]["quality_negative_generation"])
    assert not bool(result.iloc[0]["quality_train_eligible"])
    assert bool(result.iloc[1]["quality_daylight_zero"])
    assert pd.isna(result.iloc[2]["humidity_pct"])
    assert bool(result.iloc[2]["quality_invalid_weather"])
    assert result.iloc[2:]["quality_flatline"].all()


def test_quality_profiler_reports_risk_without_claiming_failure():
    report = PlantQualityProfiler().profile(_quality_frame())
    assert len(report) == 1
    assert report.iloc[0]["sensor_risk"] in {"high", "review", "low"}
    assert "failure" not in report.iloc[0]["recommended_action"]


def test_quality_audit_distinguishes_pipeline_flatline_from_official_raw(tmp_path):
    timestamps = pd.date_range("2025-01-01", periods=200, freq="h")
    model = pd.DataFrame(
        {
            "timestamp": timestamps,
            "company": "kospo",
            "plant_id": "kospo:test",
            "plant": "test",
            "region": "A",
            "energy_source": "solar",
            "generation_mwh": 0.5,
            "solar_irradiance_mj_m2": 0.5,
        }
    )
    reference = model[["timestamp", "company", "plant"]].copy()
    reference["generation_mwh"] = np.linspace(0.1, 0.9, len(reference))
    reference_path = tmp_path / "official.csv"
    reference.to_csv(reference_path, index=False)
    result = QualityAuditService().run(model, tmp_path / "audit", reference_paths=[reference_path])
    report = pd.read_csv(result.report_path, encoding="utf-8-sig")
    assert result.preprocessing_artifact_plants == 1
    assert report.iloc[0]["sensor_risk"] == "pipeline_artifact"
