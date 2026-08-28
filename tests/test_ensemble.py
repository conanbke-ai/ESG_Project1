import numpy as np
import pandas as pd

from solar_forecast.ensemble.dynamic_gate import (
    fit_dynamic_gate, fit_region_blend, normalize_prediction_columns,
    predict_dynamic_hybrid, predict_region_blend,
)
from solar_forecast.ensemble.metrics import aggregate_metrics
from solar_forecast.models.xgboost import XGBoostTrainer


def _predictions():
    return pd.DataFrame({
        "timestamp": ["2025-01-01 00:00", "2025-01-01 01:00", "2025-01-01 00:00", "2025-01-01 01:00"],
        "region": ["A", "A", "B", "B"], "plant": ["p1", "p1", "p2", "p2"],
        "y_true": [1.0, 2.0, 3.0, 4.0], "xgb_pred": [1.0, 2.0, 2.0, 3.0], "cnn_pred": [0.0, 1.0, 3.0, 4.0],
    })


def test_region_blend_learns_different_model_by_region():
    frame = _predictions()
    weights = fit_region_blend(frame)
    assert weights.set_index("region").loc["A", "xgb_weight"] == 1.0
    assert weights.set_index("region").loc["B", "xgb_weight"] == 0.0
    result = predict_region_blend(frame, weights)
    assert np.allclose(result["hybrid_pred"], result["y_true"])


def test_aggregation_recomputes_region_rmse_from_rows():
    frame = _predictions().rename(columns={"xgb_pred": "y_pred"})
    metrics = aggregate_metrics(frame)
    assert set(metrics) == {"plant", "region", "national"}
    assert metrics["national"].iloc[0]["n_samples"] == 4


def test_dynamic_gate_records_context_and_reason():
    frame = _predictions()
    gate = fit_dynamic_gate(frame, min_group_samples=1)
    result = predict_dynamic_hybrid(frame, gate)
    assert set(result["gate_scope"]) == {"plant_hour"}
    assert {"decision_reason", "selected_model", "model_disagreement"}.issubset(result.columns)
    assert np.allclose(result["xgb_weight"] + result["cnn_weight"], 1.0)
    assert result["decision_reason"].str.contains("validation evidence").all()


def test_korean_source_columns_are_normalized():
    source = pd.DataFrame({
        "일시": ["2025-01-01 00:00"], "지역": ["A"], "발전구분": ["p1"],
        "합산발전량(MWh)": [1.0], "xgb_pred": [0.9], "cnn_pred": [1.1],
    })
    result = normalize_prediction_columns(source)
    assert {"timestamp", "region", "plant", "y_true", "hour"}.issubset(result.columns)


def test_xgboost_prediction_artifact_keeps_stable_alignment_key():
    context = pd.DataFrame(
        {
            "timestamp": ["2025-01-01T00:00:00"],
            "plant_id": ["company:plant"],
            "region": ["전라남도"],
            "plant": ["영암태양광"],
        }
    )

    predictions = XGBoostTrainer._prediction_frame(
        context,
        actual=np.asarray([1.0]),
        predicted=np.asarray([0.9]),
        split="test",
    )

    assert predictions.loc[0, "plant_id"] == "company:plant"
    assert predictions.loc[0, "split"] == "test"
    assert predictions.loc[0, "y_pred"] == predictions.loc[0, "xgb_pred"]
