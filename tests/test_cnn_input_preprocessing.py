"""CPU/NumPy contract checks; runnable without Torch, Optuna, or pytest."""
from __future__ import annotations

import copy
import json
import unittest

import numpy as np

from solar_forecast.models.cnn_bilstm.input_preprocessing import (
    LEGACY_STRATEGY,
    fit_input_preprocessing,
    transform_inputs,
    validate_input_preprocessing,
)


class CnnInputPreprocessingTests(unittest.TestCase):
    def setUp(self):
        self.names = ["temperature", "radiation"]
        self.values = np.array([[1, np.nan], [np.nan, 10], [5, 30], [1000, 9000]], dtype=np.float32)
        self.selected = np.array([True, True, True, False])
        self.state = fit_input_preprocessing([(self.values, self.selected)], self.names)

    def test_train_only_population_statistics_after_imputation(self):
        self.assertEqual(self.state["feature_medians"], {"temperature": 3.0, "radiation": 20.0})
        self.assertEqual(self.state["feature_means"], {"temperature": 3.0, "radiation": 20.0})
        expected_scale = np.sqrt([8 / 3, 200 / 3])
        np.testing.assert_allclose(list(self.state["feature_scales"].values()), expected_scale)
        result = transform_inputs(self.values, self.names, self.state)
        np.testing.assert_allclose(result[:3, :2].mean(axis=0), 0.0, atol=1e-7)
        np.testing.assert_allclose(result[:3, :2].std(axis=0), 1.0, atol=1e-7)
        self.assertGreater(result[3, 0], 100)
        np.testing.assert_array_equal(result[:, 2:], [[0, 1], [1, 0], [0, 0], [0, 0]])

    def test_held_out_values_do_not_change_fitted_state(self):
        changed = self.values.copy()
        changed[3] = [-1e10, np.inf]
        candidate = fit_input_preprocessing([(changed, self.selected)], self.names)
        self.assertEqual(candidate, self.state)

    def test_multiple_plants_use_row_weighted_train_statistics(self):
        other = np.array([[9, 50], [9999, 9999]], dtype=np.float32)
        state = fit_input_preprocessing([
            (self.values, self.selected), (other, np.array([True, False])),
        ], self.names)
        # Observed medians 5/30; imputed training values [1,5,5,9]/[30,10,30,50].
        self.assertEqual(state["feature_means"], {"temperature": 5.0, "radiation": 30.0})
        np.testing.assert_allclose(list(state["feature_scales"].values()), np.sqrt([8, 200]))

    def test_constant_all_missing_and_nonfinite_values(self):
        values = np.array([[7, np.nan], [7, np.inf], [7, -np.inf], [9, 2]], dtype=np.float32)
        state = fit_input_preprocessing([(values, self.selected)], self.names)
        result = transform_inputs(values, self.names, state)
        self.assertEqual(state["all_missing_training_features"], ["radiation"])
        self.assertEqual(state["feature_scales"], {"temperature": 1.0, "radiation": 1.0})
        np.testing.assert_array_equal(result[:3, :2], 0)
        np.testing.assert_array_equal(result[:3, 2:], [[0, 1], [0, 1], [0, 1]])
        self.assertTrue(np.isfinite(result).all())
        with self.assertRaisesRegex(ValueError, "no observed values"):
            fit_input_preprocessing([(values, self.selected)], self.names, append_missing_indicators=False)

    def test_json_roundtrip_replays_without_mutation(self):
        frozen = json.loads(json.dumps(self.state))
        before = copy.deepcopy(frozen)
        for read_only in (True, False):
            source = self.values.copy()
            source.setflags(write=not read_only)
            result = transform_inputs(source, self.names, frozen)
            np.testing.assert_array_equal(source, self.values)
            self.assertFalse(np.shares_memory(result, source))
            np.testing.assert_array_equal(result, transform_inputs(self.values, self.names, self.state))
        self.assertEqual(frozen, before)

    def test_old_median_only_artifact_stays_unscaled(self):
        old = {
            "strategy": LEGACY_STRATEGY,
            "feature_medians": self.state["feature_medians"],
            "append_missing_indicators": True,
            "effective_feature_columns": self.state["effective_feature_columns"],
        }
        recognized = validate_input_preprocessing(old, self.names)
        self.assertEqual(recognized["compatibility_mode"], "legacy_median_only_no_scaling")
        np.testing.assert_array_equal(transform_inputs(self.values, self.names, old)[:, :2],
                                      [[1, 20], [3, 10], [5, 30], [1000, 9000]])
        self.assertNotIn("schema_version", old)

    def test_unknown_or_incompatible_artifacts_are_rejected(self):
        bad_states = []
        for key, value in (
            ("schema_version", 99), ("contract", "unknown"),
            ("strategy", "fit_on_test"), ("feature_columns", self.names[::-1]),
            ("effective_feature_columns", self.names), ("append_missing_indicators", "true"),
            ("missing_indicators_scaled", True), ("target_transform", "log1p"),
            ("feature_scales", {"temperature": 0, "radiation": 1}),
            ("feature_means", {"temperature": np.nan, "radiation": 1}),
            ("feature_medians", {"temperature": 1}),
        ):
            bad = copy.deepcopy(self.state)
            bad[key] = value
            bad_states.append(bad)
        for bad in bad_states:
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                transform_inputs(self.values, self.names, bad)
        with self.assertRaisesRegex(ValueError, "order"):
            transform_inputs(self.values, self.names[::-1], self.state)

    def test_invalid_feature_schema_and_missing_training_rows(self):
        with self.assertRaisesRegex(ValueError, "unique"):
            fit_input_preprocessing([(self.values, self.selected)], ["x", "x"])
        with self.assertRaisesRegex(ValueError, "collide"):
            fit_input_preprocessing([(self.values, self.selected)], ["x", "x__missing"])
        with self.assertRaisesRegex(ValueError, "No CNN training"):
            fit_input_preprocessing([(self.values, np.zeros(4, dtype=bool))], self.names)


if __name__ == "__main__":
    unittest.main()
