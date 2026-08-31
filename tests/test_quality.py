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


def _single_bucket_solar(days: int, *, hour: int = 23) -> pd.DataFrame:
    timestamps = pd.date_range("2025-01-01", periods=days * 24, freq="h")
    generation = np.zeros(len(timestamps))
    generation[timestamps.hour == hour] = 0.5
    return pd.DataFrame(
        {
            "timestamp": timestamps,
            "company": "koen",
            "plant_id": "koen:daily-total",
            "plant": "daily-total",
            "region": "A",
            "energy_source": "solar",
            "capacity_mw": 1.0,
            "generation_mwh": generation,
            "solar_irradiance_mj_m2": np.where(
                timestamps.hour == 12, 1.0, 0.0
            ),
        }
    )


def test_quality_policy_excludes_daily_total_placed_in_one_night_bucket():
    result = GenerationQualityPolicy().apply(_single_bucket_solar(40))

    assert result["quality_daily_aggregate_profile"].all()
    assert not result["quality_train_eligible"].any()
    assert result["quality_review_required"].all()
    assert result["quality_code"].str.contains("daily_aggregate_profile").all()

    report = PlantQualityProfiler().profile(_single_bucket_solar(40))
    assert bool(report.iloc[0]["daily_aggregate_profile_detected"])
    assert report.iloc[0]["aggregate_active_days"] == 40
    assert report.iloc[0]["aggregate_single_positive_day_ratio"] == 1.0
    assert report.iloc[0]["aggregate_dominant_hour"] == 23
    assert report.iloc[0]["aggregate_dominant_hour_ratio"] == 1.0
    assert report.iloc[0]["sensor_risk"] == "high"
    assert "exclude plant from hourly training" in report.iloc[0]["recommended_action"]


def test_quality_policy_does_not_flag_single_daylight_bucket_as_daily_aggregate():
    result = GenerationQualityPolicy().apply(_single_bucket_solar(40, hour=12))

    assert not result["quality_daily_aggregate_profile"].any()
    assert result["quality_train_eligible"].all()


def test_daily_aggregate_gate_does_not_cross_energy_sources_with_same_plant_id():
    solar = _single_bucket_solar(40)
    wind = solar.copy()
    wind["energy_source"] = "wind"
    result = GenerationQualityPolicy().apply(
        pd.concat([solar, wind], ignore_index=True)
    )

    assert result.loc[result["energy_source"].eq("solar"), "quality_daily_aggregate_profile"].all()
    assert not result.loc[
        result["energy_source"].eq("wind"), "quality_daily_aggregate_profile"
    ].any()


def test_quality_policy_requires_enough_consistent_active_days_for_aggregate_gate():
    too_short = GenerationQualityPolicy().apply(_single_bucket_solar(29))
    assert not too_short["quality_daily_aggregate_profile"].any()

    inconsistent = _single_bucket_solar(40)
    timestamps = pd.to_datetime(inconsistent["timestamp"])
    for day in range(3):
        target_date = timestamps.dt.normalize().drop_duplicates().iloc[day]
        same_day = timestamps.dt.normalize().eq(target_date)
        inconsistent.loc[same_day, "generation_mwh"] = 0.0
        inconsistent.loc[same_day & timestamps.dt.hour.eq(22), "generation_mwh"] = 0.5
    result = GenerationQualityPolicy().apply(inconsistent)
    assert not result["quality_daily_aggregate_profile"].any()


def test_quality_manifest_records_daily_aggregate_hard_gate(tmp_path):
    result = QualityAuditService().run(
        _single_bucket_solar(40),
        tmp_path / "audit",
    )
    manifest = __import__("json").loads(result.manifest_path.read_text(encoding="utf-8"))
    policy = manifest["policy"]["daily_aggregate_profile"]

    assert policy["minimum_active_days"] == 30
    assert policy["minimum_single_positive_day_ratio"] == 0.95
    assert policy["minimum_dominant_night_hour_ratio"] == 0.95
    assert policy["action"] == "exclude_entire_plant_from_hourly_training"
    assert manifest["daily_aggregate_profile_plants"] == {
        "count": 1,
        "plant_ids": ["koen:daily-total"],
    }
