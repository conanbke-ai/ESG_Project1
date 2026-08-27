from datetime import date

import pytest

from solar_forecast.collectors.config import CollectionConfig, load_source_catalog
from solar_forecast.collectors.kma import KmaAsosHourlyCollector
from solar_forecast.collectors.koen_browser import KoenBrowserDownloader
from solar_forecast.collectors.normalization import (
    EWP_POINT_SCHEMA,
    DailyWideGenerationNormalizer,
    EwpTrainingNormalizer,
    GENERATION_COLUMNS,
    IWEST_RENEWABLE_SCHEMA,
    IWEST_WIDE_SCHEMA,
    KOSPO_WIDE_SCHEMA,
    KoenGenerationNormalizer,
)


def test_collection_config_rejects_reverse_date_range():
    with pytest.raises(ValueError):
        CollectionConfig(start_date=date(2025, 2, 1), end_date=date(2025, 1, 1))


def test_source_catalog_uses_the_four_given_official_homepages():
    catalog = load_source_catalog()
    assert catalog["ewp"]["homepage"] == "https://www.ewp.co.kr/kor/main/"
    assert catalog["kospo"]["homepage"] == "https://www.kospo.co.kr/sites/kospo/index.do"
    assert catalog["iwest"]["homepage"] == "https://www.iwest.co.kr/sites/iwest/index.do"
    assert catalog["koen"]["homepage"] == "https://www.koenergy.kr/kosep/fr/main.do"


def test_kma_hourly_chunks_never_exceed_one_year():
    chunks = list(KmaAsosHourlyCollector._chunks(date(2023, 6, 1), date(2025, 6, 1)))
    assert chunks[0] == (date(2023, 6, 1), date(2023, 12, 31))
    assert chunks[-1][1] == date(2025, 6, 1)


def test_kma_chunks_support_leap_day():
    chunks = list(KmaAsosHourlyCollector._chunks(date(2024, 2, 29), date(2025, 3, 1)))
    assert chunks[0] == (date(2024, 2, 29), date(2024, 12, 31))


def test_koen_rejects_invalid_month_range(tmp_path):
    downloader = KoenBrowserDownloader(tmp_path)
    with pytest.raises(ValueError):
        downloader.download_year(2025, start_month=8, end_month=7)


def test_koen_normalizer_converts_24_hour_columns_to_rows():
    row = {"발전구분": "A", "호기": 1, "일자": "2025-01-01"}
    row.update({f"{hour}시 발전량(MWh)": hour for hour in range(1, 25)})
    result = KoenGenerationNormalizer().transform(__import__("pandas").DataFrame([row]))
    assert list(result.columns) == GENERATION_COLUMNS
    assert len(result) == 24
    assert result.iloc[0]["timestamp"].hour == 0
    assert result.iloc[-1]["timestamp"].hour == 23


def test_koen_normalizer_converts_downloaded_kwh_unit_to_mwh():
    row = {"발전구분": "A", "호기": 1, "일자": "2025-01-01"}
    row.update({f"{hour}시 발전량(KWh)": 1_000 for hour in range(1, 25)})
    result = KoenGenerationNormalizer().transform(__import__("pandas").DataFrame([row]))
    assert result["generation_mwh"].eq(1.0).all()


def test_koen_normalizer_corrects_legacy_mislabeled_mwh_values():
    row = {"발전구분": "A", "호기": 1, "일자": "2025-01-01"}
    row.update({f"{hour}시 발전량(MWh)": 35_000 for hour in range(1, 25)})
    result = KoenGenerationNormalizer().transform(__import__("pandas").DataFrame([row]))
    assert result["generation_mwh"].eq(35.0).all()


def test_ewp_normalizer_keeps_latest_duplicate_and_builds_time_features():
    pandas = __import__("pandas")
    base = {
        "시도명": "서울",
        "설비용량(MW)": 10,
        "발전일자": "2024-01-02 07:00",
        "기온": 5,
        "강우량(mm)": 0,
        "습도": 50,
        "적설량(mm)": 0,
        "풍속": 2,
        "적운량(10분위)": 1,
        "적운량(3분위)": 1,
        "일조(hr)": 0,
        "대기권밖일사량계산값": 0,
        "일사량": 0,
        "발전량(MWh)": 0.22,
        "윤년여부": "Y",
    }
    corrected = {**base, "발전량(MWh)": 0.2227}
    result = EwpTrainingNormalizer().transform(pandas.DataFrame([base, corrected]))
    assert len(result) == 1
    assert result.iloc[0]["generation_mwh"] == pytest.approx(0.2227)
    assert result.iloc[0]["hour"] == 7
    assert result.iloc[0]["dayofweek"] == 1


@pytest.mark.parametrize(
    ("schema", "date_column", "plant_columns", "hour_name", "source_value", "expected_mwh"),
    [
        (KOSPO_WIDE_SCHEMA, "거래일자", {"발전소명": "A", "발전기명": "B", "계량구분": "KPX"}, "1시", 1_000, 1.0),
        (
            IWEST_WIDE_SCHEMA,
            "년월일",
            {"발전기명": "A", "설비용량(MW)": 2.5},
            "01시",
            1_000_000,
            1.0,
        ),
    ],
)
def test_daily_wide_normalizer_converts_source_units(
    schema, date_column, plant_columns, hour_name, source_value, expected_mwh
):
    pandas = __import__("pandas")
    row = {date_column: "2025-01-01", **plant_columns}
    row.update({column: source_value for column in schema.hour_columns})
    result = DailyWideGenerationNormalizer(schema).transform(pandas.DataFrame([row]))
    assert len(result) == 24
    assert result.iloc[0]["generation_mwh"] == pytest.approx(expected_mwh)


def test_daily_wide_normalizer_keeps_static_capacity_and_global_plant_id():
    pandas = __import__("pandas")
    row = {"년월일": "2025-01-01", "발전기명": "A", "설비용량(MW)": 2.5}
    row.update({column: 0 for column in IWEST_WIDE_SCHEMA.hour_columns})
    result = DailyWideGenerationNormalizer(IWEST_WIDE_SCHEMA).transform(pandas.DataFrame([row]))
    assert result.iloc[0]["plant_id"] == "iwest:A"
    assert result.iloc[0]["capacity_mw"] == pytest.approx(2.5)


def test_iwest_renewable_standardization_retains_wind_with_energy_source():
    pandas = __import__("pandas")
    rows = []
    for plant in ("세종태양광", "화순풍력발전", "행원소수력"):
        row = {"날짜": "2025-01-01", "발전기명": plant, "용량_메가와트": 2.5}
        row.update({column: 1_000 for column in IWEST_RENEWABLE_SCHEMA.hour_columns})
        rows.append(row)
    result = DailyWideGenerationNormalizer(IWEST_RENEWABLE_SCHEMA).transform(pandas.DataFrame(rows))
    assert len(result) == 72
    assert set(result["energy_source"]) == {"solar", "wind", "hydro"}


def test_public_coordinate_columns_are_corrected_only_for_korea_range_inversion():
    pandas = __import__("pandas")
    row = {
        "날짜": "2025-01-01",
        "발전기명": "동해태양광",
        "설비용량(메가와트)": 1.0,
        "위도": 129.1453,
        "경도": 37.48313,
    }
    row.update({column: 1_000 for column in EWP_POINT_SCHEMA.hour_columns})
    result = DailyWideGenerationNormalizer(EWP_POINT_SCHEMA).transform(pandas.DataFrame([row]))
    assert result.iloc[0]["latitude"] == pytest.approx(37.48313)
    assert result.iloc[0]["longitude"] == pytest.approx(129.1453)
