import numpy as np
import pandas as pd
import pytest

from solar_forecast.evaluation.forecast_samples import (
    build_forecast_samples,
    forecast_evaluation_contract,
    forecast_window_positions,
)


def observations(hours=100):
    return pd.DataFrame({
        "timestamp": list(pd.date_range("2025-01-01", periods=hours, freq="h")) * 2,
        "plant_id": ["a"] * hours + ["b"] * hours,
        "temperature": np.arange(hours * 2, dtype=float),
        "generation": np.arange(hours * 2, dtype=float) + 1000,
    })


@pytest.mark.parametrize("horizon", [1, 6, 24])
def test_forecast_uses_exact_origin_weather_and_target_generation_per_plant(horizon):
    source = observations()
    samples = build_forecast_samples(source, ["temperature"], "generation", horizon)
    assert len(samples) == 2 * (100 - horizon)
    assert (samples["timestamp"] - samples["forecast_origin"]).eq(pd.Timedelta(hours=horizon)).all()
    np.testing.assert_array_equal(samples["generation"] - samples["temperature"], 1000 + horizon)
    np.testing.assert_array_equal(samples["persistence_pred"] - samples["temperature"], 1000)
    assert samples["horizon_hours"].eq(horizon).all()


def test_missing_origin_is_excluded_without_row_shift_or_cross_plant_fill():
    source = observations(20)
    missing_origin = source.loc[3, "timestamp"]
    source = source.drop(index=3)
    samples = build_forecast_samples(source, ["temperature"], "generation", 6)
    matching = samples[samples["timestamp"].eq(missing_origin + pd.Timedelta(hours=6))]
    assert matching["plant_id"].tolist() == ["b"]


def test_future_weather_mutation_cannot_change_an_earlier_forecast_input():
    source = observations()
    target_time = source.loc[50, "timestamp"]
    first = build_forecast_samples(source, ["temperature"], "generation", 24)
    source.loc[source["timestamp"].gt(target_time - pd.Timedelta(hours=24)), "temperature"] = 1e9
    second = build_forecast_samples(source, ["temperature"], "generation", 24)
    key = first["timestamp"].eq(target_time)
    np.testing.assert_array_equal(first.loc[key, "temperature"], second.loc[key, "temperature"])


@pytest.mark.parametrize("horizon", [0, -1, True, 1.5, "24", None])
def test_invalid_horizon_is_rejected(horizon):
    with pytest.raises(ValueError, match="positive integer"):
        build_forecast_samples(observations(), ["temperature"], "generation", horizon)


def test_duplicate_plant_hour_is_rejected():
    source = observations()
    with pytest.raises(ValueError, match="Duplicate plant-hour"):
        build_forecast_samples(pd.concat([source, source.iloc[:1]]), ["temperature"], "generation", 1)


def test_cnn_window_ends_at_origin_inclusive_and_excludes_gapped_history():
    times = pd.date_range("2025-01-01", periods=24, freq="h").delete(5)
    targets, origins = forecast_window_positions(times, horizon_hours=6, sequence_length=4)
    for target, origin in zip(targets, origins):
        assert times[target] - times[origin] == pd.Timedelta(hours=6)
        window = times[origin - 3:origin + 1]
        assert len(window) == 4
        assert (window[1:] - window[:-1] == pd.Timedelta(hours=1)).all()
    # Origins 06:00..08:00 need the missing 05:00 input and must be excluded.
    assert not set(times[origins].hour).intersection({6, 7, 8})
    assert {4, 9}.issubset(set(times[origins].hour))


def test_contract_distinguishes_real_horizon_from_legacy_row_estimation():
    historical = forecast_evaluation_contract("historical_forecast", 48, legacy_task="unused")
    assert historical["horizon_hours"] == 48
    assert historical["prediction_key"] == ["timestamp", "plant_id", "forecast_origin", "horizon_hours"]
    legacy = forecast_evaluation_contract(None, 24, legacy_task="observed_conditions_estimation")
    assert legacy["horizon_hours"] is None
    assert legacy["task"] == "observed_conditions_estimation"
