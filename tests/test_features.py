import numpy as np
import pandas as pd
import json
from pathlib import Path

from solar_forecast.features.engineering import SELECTED_MODEL_FEATURES, LeakageSafeFeatureEngineer


def _hourly_frame(hours: int = 220) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "timestamp": pd.date_range("2025-01-01", periods=hours, freq="h"),
            "plant_id": "kospo:test",
            "generation_mwh": np.arange(hours, dtype=float),
        }
    )


def test_history_features_use_only_values_available_24_hours_before_issue_time():
    source = _hourly_frame()
    engineer = LeakageSafeFeatureEngineer()
    original = engineer.transform(source)
    changed = source.copy()
    changed.loc[changed.index >= 190, "generation_mwh"] = 999_999
    changed_features = engineer.transform(changed)

    at_issue = original[original["timestamp"] == source.loc[190, "timestamp"]].iloc[0]
    changed_at_issue = changed_features[
        changed_features["timestamp"] == source.loc[190, "timestamp"]
    ].iloc[0]
    assert at_issue["generation_lag_24h_mwh"] == 166
    assert at_issue["generation_lag_168h_mwh"] == 22
    assert at_issue["generation_rolling_7d_mean_mwh"] == changed_at_issue[
        "generation_rolling_7d_mean_mwh"
    ]


def test_lag_is_timestamp_based_when_an_hour_is_missing():
    source = _hourly_frame(60).drop(index=26).reset_index(drop=True)
    result = LeakageSafeFeatureEngineer().transform(source)
    row = result[result["timestamp"] == pd.Timestamp("2025-01-03 02:00:00")].iloc[0]
    assert pd.isna(row["generation_lag_24h_mwh"])
    assert row["generation_lag_24h_observed"] == 0
    assert 0 <= row["generation_rolling_7d_observation_ratio"] <= 1


def test_history_engineering_rejects_duplicate_entity_timestamps():
    source = pd.concat([_hourly_frame(2), _hourly_frame(2).iloc[[0]]], ignore_index=True)
    try:
        LeakageSafeFeatureEngineer().transform(source)
    except ValueError as exc:
        assert "unique timestamp/entity" in str(exc)
    else:
        raise AssertionError("duplicate keys must be rejected")


def test_model_configs_match_the_selected_generated_columns():
    for name in ("xgboost", "cnn_bilstm"):
        values = json.loads(Path(f"config/models/{name}.json").read_text(encoding="utf-8"))
        assert values["input_dataset"] == "file/standardized/model_ready_parts"
        assert values["target_column"] == "generation_mwh"
        assert values["feature_columns"] == SELECTED_MODEL_FEATURES


def test_controlled_experiment_forbids_unavailable_subday_lags():
    values = json.loads(Path("config/experiments/controlled.json").read_text(encoding="utf-8"))
    assert values["forbidden_history_features"] == ["lag_1h", "lag_3h", "lag_6h"]
