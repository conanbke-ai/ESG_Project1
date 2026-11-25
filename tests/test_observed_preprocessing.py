"""Observed targets and missing-weather reasons survive feature construction."""
from __future__ import annotations

import unittest

import numpy as np
import pandas as pd

from solar_forecast.features.history_features import LeakageSafeFeatureEngineer
from solar_forecast.quality.generation_quality import GenerationQualityPolicy


class ObservedPreprocessingTests(unittest.TestCase):
    def frame(self):
        return pd.DataFrame({
            "timestamp": pd.date_range("2025-01-01", periods=240, freq="h"),
            "plant_id": "solar:test", "energy_source": "solar",
            "generation_mwh": 0.5, "capacity_mw": 0.1,
            "temperature_c": 10.0, "wind_speed_mps": 1.0, "humidity_pct": 50.0,
            "precipitation_mm": np.nan, "sunshine_hours": np.nan,
            "solar_irradiance_mj_m2": np.nan,
            "station_latitude": 37.5, "station_longitude": 127.0,
        })

    def test_invalid_targets_cannot_reenter_through_history(self):
        frame = self.frame()
        frame.loc[10:12, "generation_mwh"] = [-1.0, np.inf, -np.inf]
        frame.loc[13, "generation_mwh"] = 0.0
        audited = GenerationQualityPolicy().apply(frame)
        output = LeakageSafeFeatureEngineer().transform(audited)
        self.assertEqual(output.loc[10, "generation_mwh"], -1.0)
        self.assertTrue(np.isinf(output.loc[11, "generation_mwh"]))
        self.assertFalse(output.loc[10:12, "quality_train_eligible"].any())
        self.assertTrue(output.loc[34:36, "generation_lag_24h_mwh"].isna().all())
        self.assertTrue(output.loc[178:180, "generation_lag_168h_mwh"].isna().all())
        self.assertEqual(output.loc[37, "generation_lag_24h_mwh"], 0.0)
        # With 24 required valid observations, 30 available history hours minus
        # three invalid ones still give a finite mean of the retained targets.
        self.assertAlmostEqual(output.loc[53, "generation_rolling_7d_mean_mwh"], 13.0 / 27)

    def test_review_flags_do_not_delete_or_downweight_generation(self):
        frame = self.frame()
        frame["solar_irradiance_mj_m2"] = 0.5
        output = LeakageSafeFeatureEngineer().transform(GenerationQualityPolicy().apply(frame))
        self.assertTrue(output["quality_capacity_exceeded"].all())
        self.assertTrue(output["quality_flatline"].all())
        self.assertTrue(output["quality_train_eligible"].all())
        self.assertEqual(output.loc[24, "generation_lag_24h_mwh"], 0.5)

    def test_missing_weather_is_not_claimed_as_measured_zero(self):
        frame = self.frame()
        frame.loc[0, ["precipitation_mm", "sunshine_hours", "solar_irradiance_mj_m2"]] = 0.0
        frame.loc[1, "solar_irradiance_mj_m2"] = 0.02
        output = LeakageSafeFeatureEngineer().transform(GenerationQualityPolicy().apply(frame))
        self.assertEqual(output.loc[0, "precipitation_mm"], 0.0)
        self.assertEqual(output.loc[0, "precipitation_observed"], 1)
        self.assertTrue(output.loc[1:, "precipitation_mm"].isna().all())
        self.assertTrue(output.loc[1:, "sunshine_hours"].isna().all())
        self.assertEqual(output.loc[1, "solar_irradiance_mj_m2"], 0.02)
        self.assertEqual(output.loc[2, "solar_irradiance_observed"], 0)

    def test_invalid_weather_is_missing_before_observation_masks(self):
        frame = self.frame()
        frame.loc[0, "solar_irradiance_mj_m2"] = np.inf
        frame.loc[1, "humidity_pct"] = 101.0
        frame["temperature_c_invalid"] = False
        frame.loc[2, "temperature_c_invalid"] = True
        frame.loc[2, "temperature_c"] = np.nan
        output = LeakageSafeFeatureEngineer().transform(GenerationQualityPolicy().apply(frame))
        self.assertTrue(output.loc[:2, "quality_invalid_weather"].all())
        self.assertTrue(output.loc[:2, "quality_missing_weather"].all())
        self.assertEqual(output.loc[0, "solar_irradiance_observed"], 0)

    def test_later_changes_do_not_change_past_features(self):
        frame = self.frame()
        first = LeakageSafeFeatureEngineer().transform(GenerationQualityPolicy().apply(frame))
        changed = frame.copy()
        changed.loc[150:, "generation_mwh"] = -100.0
        second = LeakageSafeFeatureEngineer().transform(GenerationQualityPolicy().apply(changed))
        columns = ["generation_lag_24h_mwh", "generation_lag_168h_mwh", "generation_rolling_7d_mean_mwh"]
        pd.testing.assert_frame_equal(first.loc[:149, columns], second.loc[:149, columns])

    def test_future_daily_profile_detection_does_not_rewrite_past_history(self):
        times = pd.date_range("2024-01-01", periods=30 * 24, freq="h")
        frame = pd.DataFrame({
            "timestamp": times, "plant_id": "solar:daily-profile",
            "energy_source": "solar",
            "generation_mwh": np.where(times.hour == 2, 10.0, 0.0),
        })
        policy = GenerationQualityPolicy()
        engineer = LeakageSafeFeatureEngineer()
        first = engineer.transform(policy.apply(frame.iloc[:29 * 24]))
        second = engineer.transform(policy.apply(frame))
        # The existing retrospective profiler reaches its 30-day threshold.
        # Its whole-plant label must never change earlier lag/rolling inputs.
        self.assertFalse(first["quality_daily_aggregate_profile"].any())
        self.assertTrue(second["quality_daily_aggregate_profile"].all())
        columns = [
            "generation_lag_24h_mwh", "generation_lag_168h_mwh",
            "generation_rolling_7d_mean_mwh", "generation_lag_24h_observed",
            "generation_lag_168h_observed", "generation_rolling_7d_observation_ratio",
        ]
        pd.testing.assert_frame_equal(first[columns], second.iloc[:len(first)][columns])
        self.assertEqual(second.loc[26, "generation_lag_24h_mwh"], 10.0)

    def test_unknown_solar_position_is_not_classified_as_night(self):
        frame = self.frame()
        frame.loc[12, "station_latitude"] = np.nan
        frame.loc[13, "station_longitude"] = np.nan
        output = LeakageSafeFeatureEngineer().transform(frame)
        self.assertTrue(output.loc[12:13, "solar_elevation_sin"].isna().all())
        self.assertTrue(output.loc[12:13, "is_daylight"].isna().all())
        self.assertEqual(output.loc[0, "is_daylight"], 0.0)
        self.assertEqual(output.loc[36, "is_daylight"], 1.0)


if __name__ == "__main__":
    unittest.main()
