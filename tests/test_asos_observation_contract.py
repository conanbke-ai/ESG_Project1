"""ASOS station-hour normalization and provenance regression tests."""
from pathlib import Path
import tempfile
import unittest

import numpy as np
import pandas as pd

from solar_forecast.datasets.asos_weather_store import (
    merge_asos_observations, merge_weather_rows, observation_to_weather_row,
)
from solar_forecast.features.asos_features import (
    KmaAsosNormalizer, WEATHER_PROVENANCE_COLUMNS,
)


def observations(times, **overrides):
    frame = pd.DataFrame({
        "지점": 108, "지점명": "서울", "일시": times,
        "기온(°C)": 20.0, "강수량(mm)": np.nan,
        "풍속(m/s)": 2.0, "습도(%)": 50.0,
        "일조(hr)": np.nan, "일사(MJ/m2)": np.nan,
        "전운량(10분위)": 5.0, "중하층운량(10분위)": 3.0,
    })
    for column, values in overrides.items():
        frame[column] = values
    return frame


def station(start="2020-01-01", end=None, latitude=37.0, elevation=80.0):
    return {"지점": 108, "시작일": start, "종료일": end,
            "위도": latitude, "경도": 127.0, "노장해발고도(m)": elevation}


class AsosObservationContractTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.root = Path(self.directory.name)
        self.metadata = self.root / "metadata.csv"
        self.write_metadata([station()])
        self.normalizer = KmaAsosNormalizer(self.metadata)

    def tearDown(self):
        self.directory.cleanup()

    def write_metadata(self, records):
        pd.DataFrame(records).to_csv(self.metadata, index=False)

    def test_preserves_blank_zero_and_qc_reasons_without_mutating_input(self):
        source = observations(pd.date_range("2021-01-01", periods=4, freq="h"), **{
            "기온(°C)": [20, 999, -5, 21], "taQcflg": ["", "1", "0", "2"],
            "강수량(mm)": [np.nan, 0, 1, 3], "rnQcflg": ["", "0", "9", ""],
        })
        before = source.copy(deep=True)
        result = self.normalizer.transform(source)
        pd.testing.assert_frame_equal(source, before)
        self.assertTrue(set(WEATHER_PROVENANCE_COLUMNS).issubset(result.columns))
        self.assertEqual(result["temperature_c_missing_reason"].tolist(),
                         ["observed", "qc_error", "observed", "observed"])
        self.assertEqual(result["precipitation_mm_missing_reason"].tolist(),
                         ["source_missing", "observed", "qc_missing", "observed"])
        self.assertTrue(pd.isna(result.loc[0, "precipitation_mm"]))
        self.assertEqual(result.loc[1, "precipitation_mm"], 0)
        self.assertTrue(pd.isna(result.loc[2, "precipitation_mm"]))
        self.assertEqual(result.loc[3, "temperature_c_qc_flag"], "2")
        self.assertTrue(result["weather_station_hour_present"].all())
        # Nighttime blanks remain unknown; observed negative temperature stays.
        self.assertTrue(result["sunshine_hours"].isna().all())
        self.assertEqual(result.loc[2, "temperature_c"], -5)

    def test_non_finite_non_numeric_and_ranges_are_masked_with_reason(self):
        source = observations(pd.date_range("2021-01-01", periods=4, freq="h"), **{
            "기온(°C)": [np.inf, -999, "bad", 0],
            "풍속(m/s)": [np.inf, -1, 2, 0],
            "일조(hr)": [0, 1.1, -1, 1],
            "습도(%)": [101, -1, 100, 0],
        })
        result = self.normalizer.transform(source)
        self.assertEqual(result["temperature_c_missing_reason"].tolist(),
                         ["non_finite", "out_of_range", "non_numeric", "observed"])
        self.assertTrue(result.loc[:2, "temperature_c_invalid"].all())
        self.assertFalse(result.loc[3, "temperature_c_invalid"])
        self.assertTrue(result.loc[:1, "humidity_pct"].isna().all())
        self.assertEqual(result.loc[3, "sunshine_hours"], 1)
        self.assertFalse(np.isinf(result.select_dtypes(include="number")).any().any())

    def test_metadata_move_boundary_and_expiry_do_not_borrow_latest(self):
        self.write_metadata([
            station("2020-01-01", "2021-06-30", latitude=37, elevation=80),
            station("2021-06-30", "2022-01-01", latitude=38, elevation=90),
        ])
        result = self.normalizer.transform(observations([
            "2019-12-31 23:00", "2021-06-29 23:00", "2021-06-30 00:00",
            "2022-01-01 23:00", "2022-01-02 00:00",
        ]))
        self.assertEqual(result["station_metadata_status"].tolist(),
                         ["outside_validity", "matched", "matched", "matched", "outside_validity"])
        self.assertTrue(pd.isna(result.loc[0, "station_latitude"]))
        self.assertEqual(result.loc[1, "station_latitude"], 37)
        self.assertEqual(result.loc[2, "station_latitude"], 38)
        self.assertEqual(result.loc[2, "station_elevation_m"], 90)
        self.assertTrue(pd.isna(result.loc[4, "station_latitude"]))

    def test_future_observations_and_metadata_do_not_modify_past(self):
        source = observations(["2021-01-01 00:00", "2021-01-01 01:00"])
        before = self.normalizer.transform(source)
        self.write_metadata([station(), station("2026-01-01", latitude=38, elevation=99)])
        expanded = pd.concat([source, observations(["2026-01-01 00:00"], **{
            "기온(°C)": 50, "강수량(mm)": 100,
        })], ignore_index=True)
        after = self.normalizer.transform(expanded).iloc[:len(source)].reset_index(drop=True)
        pd.testing.assert_frame_equal(before, after)

    def test_absent_and_undated_metadata_have_explicit_status(self):
        record = station()
        del record["시작일"]
        del record["종료일"]
        self.write_metadata([record])
        result = self.normalizer.transform(observations(["2021-01-01", "2021-01-02"], **{
            "지점": [108, 999],
        }))
        self.assertEqual(result["station_metadata_status"].tolist(),
                         ["validity_unknown", "station_not_found"])
        self.assertEqual(result.loc[0, "station_latitude"], 37)
        self.assertTrue(pd.isna(result.loc[1, "station_latitude"]))

    def test_malformed_metadata_does_not_become_open_ended_history(self):
        self.write_metadata([station("2020-01-01", "unknown")])
        result = self.normalizer.transform(observations(["2021-01-01"]))
        self.assertEqual(result.loc[0, "station_metadata_status"], "validity_unknown")
        self.assertTrue(pd.isna(result.loc[0, "station_latitude"]))

    def test_shared_station_and_latest_duplicate_have_identical_provenance(self):
        raw = observations(["2021-01-01 00:00", "2021-01-01 00:00"], **{
            "기온(°C)": [20, 21], "taQcflg": ["1", "0"],
        })
        weather = self.normalizer.transform(raw)
        self.assertEqual(len(weather), 1)
        plants = pd.DataFrame({"plant_id": ["a", "b"], "station_id": [108, 108],
                               "timestamp": pd.to_datetime(["2021-01-01"] * 2)})
        joined = plants.merge(weather, on=["station_id", "timestamp"], validate="many_to_one")
        self.assertEqual(joined["temperature_c"].tolist(), [21, 21])
        pd.testing.assert_series_equal(joined[WEATHER_PROVENANCE_COLUMNS].iloc[0],
                                       joined[WEATHER_PROVENANCE_COLUMNS].iloc[1], check_names=False)

    def test_api_masks_only_documented_bad_flags_and_clears_stale_qc(self):
        base = {"stnId": "108", "tm": "2021-01-01 00:00", "ta": "20"}
        self.assertEqual(observation_to_weather_row({**base, "taQcflg": "2"})["기온(°C)"], "20")
        self.assertEqual(observation_to_weather_row({**base, "taQcflg": "9.0"})["기온(°C)"], "")
        merge_asos_observations(self.root, [{**base, "taQcflg": "1"}])
        merge_asos_observations(self.root, [{**base, "ta": "21"}])
        updated = pd.read_csv(self.root / "OBS_ASOS_TIM_2021.csv", encoding="cp949")
        self.assertEqual(updated.loc[0, "기온(°C)"], 21)
        self.assertTrue(pd.isna(updated.loc[0, "taQcflg"]))

    def test_download_qc_cannot_attach_to_an_api_replacement(self):
        merge_weather_rows(self.root, [{
            "지점": "108", "일시": "2021-01-01 00:00", "기온(°C)": "999",
            "기온 QC플래그": "1", "습도(%)": "101", "hmQcflg": "1",
        }])
        merge_asos_observations(self.root, [{
            "stnId": "108", "tm": "2021-01-01 00:00", "ta": "21", "taQcflg": "0",
        }])
        updated = pd.read_csv(self.root / "OBS_ASOS_TIM_2021.csv", encoding="cp949")
        self.assertTrue(pd.isna(updated.loc[0, "기온 QC플래그"]))
        self.assertEqual(updated.loc[0, "taQcflg"], 0)
        self.assertEqual(updated.loc[0, "hmQcflg"], 1)


if __name__ == "__main__":
    unittest.main()
