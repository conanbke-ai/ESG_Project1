from __future__ import annotations

import pandas as pd
import pytest

from solar_forecast.evaluation.temporal_split import TemporalSplitConfig
from solar_forecast.evaluation.training_eligibility import audit_training_eligibility


def _frame() -> pd.DataFrame:
    times = pd.date_range("2024-01-01", periods=240, freq="h")
    rows = []
    for plant_id in ("plant-a", "plant-b"):
        for index, timestamp in enumerate(times):
            weather = float(index % 10)
            if plant_id == "plant-b" and index >= 204:
                weather = None
            rows.append(
                {
                    "timestamp": timestamp,
                    "plant_id": plant_id,
                    "generation_mwh": float(index % 24),
                    "quality_train_eligible": True,
                    "temperature_c": weather,
                    "solar_irradiance_mj_m2": weather,
                }
            )
    return pd.DataFrame(rows)


def test_eligibility_keeps_complete_overlap_and_rejects_empty_test_overlap() -> None:
    report = audit_training_eligibility(
        _frame(),
        TemporalSplitConfig(
            validation_fraction=0.20,
            calibration_fraction=0.15,
            test_fraction=0.15,
            gap_hours=0,
        ),
        weather_columns=("temperature_c", "solar_irradiance_mj_m2"),
    )

    plants = {item["plant_id"]: item for item in report["plants"]}
    assert report["fixed_start_year_used"] is False
    assert report["final_thresholds_applied"] is False
    assert report["final_training_selection_ready"] is False
    assert plants["plant-a"]["status"] == "CANDIDATE_REQUIRES_THRESHOLD_REVIEW"
    assert plants["plant-a"]["hard_reject_reasons"] == []
    assert plants["plant-b"]["status"] == "STRUCTURAL_REJECT"
    assert "empty_overlap_splits:test" in plants["plant-b"]["hard_reject_reasons"]
    assert report["population"]["candidate_plants"] == 1
    assert report["population"]["structurally_rejected_plants"] == 1


def test_eligibility_reports_hourly_gaps_without_hardcoding_duration_threshold() -> None:
    frame = _frame()
    frame = frame.loc[
        ~(
            frame["plant_id"].eq("plant-a")
            & frame["timestamp"].eq(pd.Timestamp("2024-01-03T12:00:00"))
        )
    ].reset_index(drop=True)

    report = audit_training_eligibility(
        frame,
        TemporalSplitConfig(
            validation_fraction=0.20,
            calibration_fraction=0.15,
            test_fraction=0.15,
            gap_hours=0,
        ),
        weather_columns=("temperature_c", "solar_irradiance_mj_m2"),
    )
    plant = next(item for item in report["plants"] if item["plant_id"] == "plant-a")

    assert plant["target_continuity"]["gap_count_gt_1h"] == 1
    assert plant["target_continuity"]["max_gap_hours"] == 2.0
    assert plant["target_continuity"]["hourly_coverage"] < 1.0
    assert plant["status"] == "CANDIDATE_REQUIRES_THRESHOLD_REVIEW"


def test_eligibility_rejects_duplicate_plant_hour_keys() -> None:
    frame = _frame()
    duplicate = frame.iloc[[0]].copy()
    frame = pd.concat([frame, duplicate], ignore_index=True)

    with pytest.raises(ValueError, match="Duplicate quality-eligible plant-hour keys"):
        audit_training_eligibility(
            frame,
            TemporalSplitConfig(
                validation_fraction=0.20,
                calibration_fraction=0.15,
                test_fraction=0.15,
                gap_hours=0,
            ),
            weather_columns=("temperature_c", "solar_irradiance_mj_m2"),
        )
