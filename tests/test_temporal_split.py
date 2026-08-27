import pandas as pd

from solar_forecast.evaluation.temporal import TemporalSplitConfig, TemporalSplitter


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
