from __future__ import annotations

import numpy as np

from solar_forecast.evaluation.regression_metrics import validation_diagnostics


def test_validation_diagnostics_exposes_tail_error_and_capacity_normalization():
    truth = np.array([0.0, 1.0, 1.0, 1.0, 1.0])
    prediction = np.array([0.0, 1.0, 1.0, 1.0, 6.0])
    persistence = np.array([0.0, 0.5, 0.5, 0.5, 0.5])
    daylight = np.array([False, True, True, True, True])
    capacity = np.array([10.0] * 5)
    plant_id = np.array(["a", "a", "a", "b", "b"])
    plant = np.array(["A", "A", "A", "B", "B"])
    region = np.array(["R1", "R1", "R1", "R2", "R2"])
    timestamp = np.arange(
        np.datetime64("2025-01-01T00"),
        np.datetime64("2025-01-01T05"),
        np.timedelta64(1, "h"),
    )

    result = validation_diagnostics(
        truth,
        prediction,
        persistence_pred=persistence,
        is_daylight=daylight,
        capacity_mw=capacity,
        plant_id=plant_id,
        plant=plant,
        region=region,
        timestamp=timestamp,
        top_k=3,
    )

    assert result["mae_mwh"] == 1.0
    assert result["rmse_mwh"] > result["mae_mwh"]
    assert result["abs_error_max_mwh"] == 5.0
    assert result["abs_error_p99_mwh"] > 4.0
    assert result["daylight_mae_mwh"] == 1.25
    assert result["daylight_rmse_mwh"] > result["daylight_mae_mwh"]
    assert result["nmae_capacity_pct"] == 10.0
    assert result["persistence_skill_pct"] < 0
    assert result["top_absolute_errors"][0]["plant_id"] == "b"
    assert result["top_absolute_errors"][0]["abs_error_mwh"] == 5.0
    assert len(result["plant_metrics"]) == 2
    by_id = {row["plant_id"]: row for row in result["plant_metrics"]}
    assert by_id["b"]["rmse_mwh"] > by_id["a"]["rmse_mwh"]
    assert by_id["b"]["nmae_capacity_pct"] == 25.0


def test_validation_diagnostics_handles_missing_optional_context():
    truth = np.array([1.0, 2.0, 3.0])
    prediction = np.array([1.0, 2.5, 2.5])

    result = validation_diagnostics(truth, prediction)

    assert result["rows"] == 3
    assert result["daylight_mae_mwh"] is None
    assert result["daylight_rmse_mwh"] is None
    assert result["nmae_capacity_pct"] is None
    assert result["persistence_skill_pct"] is None
    assert result["plant_metrics"] == []
    assert len(result["top_absolute_errors"]) == 3
