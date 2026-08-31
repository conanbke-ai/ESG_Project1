from datetime import date
from pathlib import Path

import pytest

from solar_forecast.collectors.config import CollectionConfig, load_source_catalog
from solar_forecast.collectors.archive import HistoricalGenerationStandardizationService
from solar_forecast.collectors.kma import KmaAsosHourlyCollector
from solar_forecast.collectors.koen_browser import KoenBrowserDownloader
from solar_forecast.collectors.csv_artifacts import inspect_csv_artifact, write_standardized_csv
from solar_forecast.collectors.naming import (
    build_solar_download_filename,
    is_canonical_solar_download_name,
)
from solar_forecast.collectors.normalization import (
    EWP_POINT_SCHEMA,
    DailyWideGenerationNormalizer,
    EwpTrainingNormalizer,
    GENERATION_COLUMNS,
    KOSPO_ARCHIVE_INTERVAL_SCHEMA,
    IWEST_RENEWABLE_SCHEMA,
    IWEST_WIDE_SCHEMA,
    KOSPO_WIDE_SCHEMA,
    KoenGenerationNormalizer,
)
from solar_forecast.collectors.openapi import KomipoRenewableCollector


def test_collection_config_rejects_reverse_date_range():
    with pytest.raises(ValueError):
        CollectionConfig(start_date=date(2025, 2, 1), end_date=date(2025, 1, 1))


def test_source_catalog_uses_the_four_given_official_homepages():
    catalog = load_source_catalog()
    assert catalog["ewp"]["homepage"] == "https://www.ewp.co.kr/kor/main/"
    assert catalog["kospo"]["homepage"] == "https://www.kospo.co.kr/sites/kospo/index.do"
    assert catalog["iwest"]["homepage"] == "https://www.iwest.co.kr/sites/iwest/index.do"
    assert catalog["koen"]["homepage"] == "https://www.koenergy.kr/kosep/fr/main.do"
    assert catalog["kospo"]["organization"] == "한국남부발전(주)"
    assert catalog["iwest"]["detail_name"] == "태양광통합"


def test_download_filename_uses_canonical_korean_rule():
    filename = build_solar_download_filename(
        "한국남부발전(주)",
        "남제주소내",
        date(2025, 2, 28),
    )
    assert filename == "한국남부발전(주)_[남제주소내] 태양광발전실적_20250228.csv"
    assert is_canonical_solar_download_name(filename)


def test_download_filename_sanitizes_path_characters():
    filename = build_solar_download_filename(
        "한국남동발전(주)",
        "월간/통합[202501]",
        date(2026, 8, 28),
    )
    assert filename == "한국남동발전(주)_[월간_통합_202501_] 태양광발전실적_20260828.csv"
    assert is_canonical_solar_download_name(filename)


def test_standardized_csv_is_atomic_utf8_sig(tmp_path):
    pandas = __import__("pandas")
    destination = tmp_path / "표준화.csv"
    write_standardized_csv(pandas.DataFrame({"발전소": ["남제주소내"]}), destination)

    audit = inspect_csv_artifact(destination)
    assert audit.encoding == "utf-8-sig"
    assert audit.byte_order_mark is True
    assert not destination.with_suffix(".csv.part").exists()
    assert pandas.read_csv(destination, encoding="utf-8-sig").iloc[0, 0] == "남제주소내"


def test_koen_bronze_move_preserves_provider_bytes(tmp_path):
    source = tmp_path / "provider.csv"
    original = "발전구분,호기\n해창만,1\n".encode("cp949")
    source.write_bytes(original)
    destination = tmp_path / "raw" / "download.csv"

    downloader = KoenBrowserDownloader(tmp_path, downloaded_on=date(2026, 8, 28))
    downloader._move_original(source, destination)

    assert destination.read_bytes() == original
    assert not source.exists()


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
    assert result.attrs["generation_unit_resolution"] == {
        "declared_source_unit": "mwh",
        "resolved_source_unit": "kwh",
        "method": "legacy_header_scale_correction",
        "capacity_factor_p99": None,
        "capacity_factor_max": None,
    }


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


def test_kospo_interval_normalizer_corrects_decisive_wh_scale_header_defect():
    pandas = __import__("pandas")
    row = {
        "년월일": "2025-01-01",
        "_source_plant": "익산 다송리",
        "호기": "A",
        "설비용량(kW)": 893,
        "총량(kWh)": 7_200_000,
    }
    row.update({column: 0 for column in KOSPO_ARCHIVE_INTERVAL_SCHEMA.hour_columns})
    for column in KOSPO_ARCHIVE_INTERVAL_SCHEMA.hour_columns[8:17]:
        row[column] = 800_000

    result = DailyWideGenerationNormalizer(KOSPO_ARCHIVE_INTERVAL_SCHEMA).transform(
        pandas.DataFrame([row]),
        source_file="한국남부발전(주)_[익산 다송리] 태양광발전실적_20250808.csv",
    )

    assert result["generation_mwh"].max() == pytest.approx(0.8)
    assert result["capacity_mw"].eq(0.893).all()
    resolution = result.attrs["generation_unit_resolution"]
    assert resolution["declared_source_unit"] == "kwh"
    assert resolution["resolved_source_unit"] == "wh"
    assert resolution["method"] == "capacity_factor_1000x_correction"
    assert resolution["capacity_factor_max"] == pytest.approx(0.8 / 0.893)


def test_kospo_interval_normalizer_preserves_physically_plausible_kwh_values():
    pandas = __import__("pandas")
    row = {
        "년월일": "2025-01-01",
        "_source_plant": "정상 태양광",
        "호기": "A",
        "설비용량(kW)": 893,
        "총량(kWh)": 19_200,
    }
    row.update({column: 800 for column in KOSPO_ARCHIVE_INTERVAL_SCHEMA.hour_columns})

    result = DailyWideGenerationNormalizer(KOSPO_ARCHIVE_INTERVAL_SCHEMA).transform(
        pandas.DataFrame([row])
    )

    assert result["generation_mwh"].eq(0.8).all()
    resolution = result.attrs["generation_unit_resolution"]
    assert resolution["resolved_source_unit"] == "kwh"
    assert resolution["method"] == "declared_header"


def test_daily_total_reconciliation_fails_closed_on_incomplete_source_row():
    pandas = __import__("pandas")
    row = {
        "년월일": "2025-01-01",
        "_source_plant": "불완전 태양광",
        "호기": "A",
        "설비용량(kW)": 893,
        "총량(kWh)": 800,
    }
    row.update({column: 0 for column in KOSPO_ARCHIVE_INTERVAL_SCHEMA.hour_columns})
    row[KOSPO_ARCHIVE_INTERVAL_SCHEMA.hour_columns[12]] = pandas.NA

    with pytest.raises(ValueError, match="requires all 24 hourly buckets"):
        DailyWideGenerationNormalizer(KOSPO_ARCHIVE_INTERVAL_SCHEMA).transform(
            pandas.DataFrame([row])
        )


def test_actual_dasong_export_reconciles_both_daily_total_unit_segments():
    source = (
        Path(__file__).parents[1]
        / "file"
        / "solar_data_file"
        / "한국남부발전"
        / "한국남부발전(주)_[익산 다송리] 태양광발전실적_20250808.csv"
    )

    result = DailyWideGenerationNormalizer(KOSPO_ARCHIVE_INTERVAL_SCHEMA).read(source)

    assert len(result) == 14_064
    assert result["generation_mwh"].max() == pytest.approx(0.813553)
    resolution = result.attrs["generation_unit_resolution"]
    assert resolution["resolved_source_unit"] == "wh"
    daily_total = resolution["daily_total"]
    assert daily_total["status"] == "mixed"
    assert daily_total["evidence_unit_counts"] == {"kwh": 258, "wh": 312}
    assert daily_total["resolved_unit_counts"] == {"kwh": 274, "wh": 312}
    assert daily_total["zero_rows"] == 16
    assert daily_total["unreconciled_rows"] == 0
    assert daily_total["resolved_unit_date_ranges"] == {
        "kwh": {"start": "2024-01-01", "end": "2024-09-30"},
        "wh": {"start": "2024-10-01", "end": "2025-08-08"},
    }


def test_generation_manifest_records_capacity_based_unit_resolution(tmp_path):
    pandas = __import__("pandas")
    source_root = tmp_path / "solar_data_file"
    source_dir = source_root / "한국남부발전"
    source_dir.mkdir(parents=True)
    source = source_dir / "한국남부발전(주)_[익산 다송리] 태양광발전실적_20250808.csv"
    row = {
        "년월일": "2025-01-01",
        "호기": "A",
        "설비용량(kW)": 893,
        "총량(kWh)": 19_200_000,
    }
    row.update({column: 800_000 for column in KOSPO_ARCHIVE_INTERVAL_SCHEMA.hour_columns})
    pandas.DataFrame([row]).to_csv(source, index=False, encoding="cp949")

    run = HistoricalGenerationStandardizationService(
        source_root,
        tmp_path / "standardized",
    ).run()

    assert len(run.partitions) == 1
    partition = run.partitions[0]
    assert partition.declared_source_unit == "kwh"
    assert partition.resolved_source_unit == "wh"
    assert partition.unit_resolution_method == "capacity_factor_1000x_correction"
    assert partition.unit_capacity_factor_max == pytest.approx(0.8 / 0.893)
    assert partition.daily_total_resolution["status"] == "consistent"
    assert partition.daily_total_resolution["resolved_unit_counts"] == {"wh": 1}
    standardized = pandas.read_csv(partition.destination, encoding="utf-8-sig")
    assert standardized["generation_mwh"].eq(0.8).all()


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


class _XmlResponse:
    def __init__(self, content: bytes):
        self.content = content

    def raise_for_status(self):
        return None


class _KomipoSession:
    def __init__(self):
        self.params = []

    def get(self, _url, *, params, timeout):
        self.params.append((params, timeout))
        return _XmlResponse(
            b"""<response><header><resultCode>00</resultCode><resultMsg>NORMAL</resultMsg></header>
            <body><totalCount>1</totalCount><items><item>
            <siteterm>Boryeong</siteterm><unitterm>Solar-1</unitterm>
            <gathdtm>2025-01-01 12:00:00</gathdtm><daypower>6287</daypower>
            </item></items></body></response>"""
        )


def test_komipo_collector_writes_resumable_bounded_bronze_partition(tmp_path, monkeypatch):
    monkeypatch.setenv("DATA_GO_SERVICE_KEY", "test-key")
    session = _KomipoSession()
    config = CollectionConfig(
        start_date=date(2025, 1, 1),
        end_date=date(2025, 1, 1),
        output_dir=tmp_path,
        komipo_station_codes=["8509"],
        api_max_calls=1,
    )
    result = KomipoRenewableCollector(config, session=session).collect()

    assert result.status == "downloaded"
    assert result.rows == 1
    assert len(session.params) == 1
    assert result.files[0] == Path(tmp_path) / "komipo/station=8509/year=2025/date=20250101.csv.gz"
    frame = __import__("pandas").read_csv(result.files[0])
    assert frame.iloc[0]["generation_value"] == 6287
    assert frame.iloc[0]["source_unit"] == "portal_unspecified"


def test_komipo_collector_preflights_daily_call_budget(tmp_path, monkeypatch):
    monkeypatch.setenv("DATA_GO_SERVICE_KEY", "test-key")
    config = CollectionConfig(
        start_date=date(2025, 1, 1),
        end_date=date(2025, 1, 2),
        output_dir=tmp_path,
        komipo_station_codes=["8509"],
        api_max_calls=1,
    )
    result = KomipoRenewableCollector(config, session=_KomipoSession()).collect()
    assert result.status == "configuration_required"
    assert "at least 2 calls" in result.message
