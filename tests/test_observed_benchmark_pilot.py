"""Keep the observed integration pilot's preflight and executed cohort identical."""
from copy import deepcopy
import json
from pathlib import Path

import pandas as pd
import pytest

from solar_forecast.config_loader import PROJECT_ROOT
from solar_forecast.evaluation.experiment_config import build_candidate_configs, load_experiment_config
from solar_forecast.evaluation.temporal_split import TemporalSplitConfig, TemporalSplitter
from tools.run_observed_benchmark_pilot import build_pilot_config, select_observed_plant_year


def test_pilot_freezes_audited_calendar_and_excludes_unavailable_forecast_archives(tmp_path):
    base = load_experiment_config(PROJECT_ROOT / "config/experiments/optimized.json")
    original = deepcopy(base)
    times = pd.date_range("2023-12-25", "2024-12-31T23:00:00", freq="h")
    frame = pd.DataFrame({
        "timestamp": times, "plant_id": "fixture", "energy_source": "solar",
        "quality_train_eligible": True, "generation_mwh": 1.0,
        "temperature_c": 20.0, "wind_speed_mps": 1.0, "humidity_pct": 50.0,
    })
    cohort, selection = select_observed_plant_year(frame, split=base["split"])
    values = build_pilot_config(base, tmp_path / "observed.csv", tmp_path, PROJECT_ROOT, {"cohort": selection})
    assert base == original
    assert values["pilot_overrides"]["requested_population_split"] == original["split"]
    assert values["pilot_overrides"]["requested_future_weather"] == original["future_weather"]
    assert values["pilot_overrides"]["model_performance_used_for_boundaries"] is False
    assert values["pilot_overrides"]["new_holdout_evidence"] is False
    assert values["split"]["test_end"] == selection["input_end"]
    assert values["split"]["purge_gap_hours"] == 168
    config_file = tmp_path / "pilot.json"
    config_file.write_text(json.dumps(values))
    load_experiment_config(config_file, project_root=PROJECT_ROOT)

    # The old production calendar has no Test in this one-year fixture.
    requested = TemporalSplitter(TemporalSplitConfig.from_mapping(original["split"]))
    assert not requested.labels(cohort.timestamp, requested.boundaries(cohort.timestamp)).eq("test").any()
    expected = selection["continuous_window_coverage"]["boundaries"]
    for horizon in values["horizons_hours"]:
        candidates = build_candidate_configs(values, horizon, tmp_path, project_root=PROJECT_ROOT)
        assert len(candidates) == 3  # one XGBoost + two CNN lookbacks
        for _, candidate in candidates:
            split = TemporalSplitter(TemporalSplitConfig.from_mapping(candidate.values))
            actual = split.boundaries(cohort.timestamp)
            assert all(actual.to_dict()[key] == value for key, value in expected.items())
            partitions = split.split_frame(cohort)
            assert all(len(getattr(partitions, name)) > 0 for name in ("train", "validation", "calibration", "test"))
            assert candidate.values["future_weather"] == {"enabled": False}
            assert candidate.values["optimizer"]["max_trials"] == (2 if candidate.model == "xgboost" else 1)


def test_pilot_rejects_unqualified_cohort(tmp_path):
    with pytest.raises(ValueError, match="qualified observed cohort"):
        build_pilot_config({}, Path("unused.csv"), tmp_path, tmp_path,
                           {"cohort": {"continuous_window_coverage": {"qualified": False}}})
