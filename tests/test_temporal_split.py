import pandas as pd

from solar_forecast.evaluation.temporal_split import TemporalSplitConfig
from solar_forecast.evaluation.temporal_split import TemporalSplitter


def test_global_timestamp_boundaries_are_shared_by_irregular_entities():
    hours = pd.date_range("2025-01-01", periods=100, freq="h")
    frame = pd.concat(
        [
            pd.DataFrame({"timestamp": hours, "plant_id": "complete"}),
            pd.DataFrame({"timestamp": hours[35:], "plant_id": "late_start"}),
        ],
        ignore_index=True,
    )
    splits = TemporalSplitter(
        TemporalSplitConfig(
            validation_fraction=0.15,
            calibration_fraction=0.10,
            test_fraction=0.15,
        )
    ).split_frame(frame)

    assert splits.train["timestamp"].max() == splits.boundaries.train_end
    assert splits.validation["timestamp"].min() > splits.boundaries.train_end
    assert splits.calibration["timestamp"].min() > splits.boundaries.validation_end
    assert splits.test["timestamp"].min() > splits.boundaries.calibration_end
    for partition in (splits.validation, splits.calibration, splits.test):
        assert set(partition["plant_id"]) == {"complete", "late_start"}


def test_purge_gap_removes_boundary_hours_from_all_entities():
    frame = pd.DataFrame({"timestamp": pd.date_range("2025-01-01", periods=200, freq="h")})
    splits = TemporalSplitter(
        TemporalSplitConfig(
            validation_fraction=0.15,
            calibration_fraction=0.10,
            test_fraction=0.15,
            gap_hours=4,
        )
    ).split_frame(frame)
    assert splits.validation["timestamp"].min() > splits.boundaries.train_end + pd.Timedelta(hours=4)
    assert splits.calibration["timestamp"].min() > splits.boundaries.validation_end + pd.Timedelta(hours=4)
    assert splits.test["timestamp"].min() > splits.boundaries.calibration_end + pd.Timedelta(hours=4)


def test_global_calendar_boundaries_keep_same_test_period_for_late_start_plant():
    hours = pd.date_range("2024-01-01", "2025-12-31 23:00:00", freq="h")
    frame = pd.concat(
        [
            pd.DataFrame({"timestamp": hours, "plant_id": "long_history"}),
            pd.DataFrame({"timestamp": hours[24 * 120 :], "plant_id": "shorter_history"}),
        ],
        ignore_index=True,
    )
    config = TemporalSplitConfig(
        train_end="2024-06-30T23:00:00",
        validation_end="2024-09-30T23:00:00",
        calibration_end="2024-12-31T23:00:00",
        test_end="2025-12-31T23:00:00",
        gap_hours=168,
    )
    splits = TemporalSplitter(config).split_frame(frame)

    assert splits.boundaries.to_dict()["split_mode"] == "calendar"
    assert splits.test["timestamp"].min() == pd.Timestamp("2025-01-08T00:00:00")
    for plant_id in ("long_history", "shorter_history"):
        plant_test = splits.test.loc[splits.test["plant_id"].eq(plant_id), "timestamp"]
        assert not plant_test.empty
        assert plant_test.min() == pd.Timestamp("2025-01-08T00:00:00")
        assert plant_test.max() == pd.Timestamp("2025-12-31T23:00:00")
