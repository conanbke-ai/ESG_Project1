from pathlib import Path

import pandas as pd

from solar_forecast.collectors.metadata import PlantMetadataCatalog
from solar_forecast.features.registry import KmaStationCatalog, parse_administrative_area
from solar_forecast.features.service import NationwideModelDatasetBuilder


def _stations() -> KmaStationCatalog:
    return KmaStationCatalog(
        pd.DataFrame(
            [
                {
                    "station_id": 266,
                    "station_name": "광양시",
                    "station_address": "전라남도 광양시 중동",
                    "station_latitude": 34.9434,
                    "station_longitude": 127.6914,
                    "admin_province": "전라남도",
                    "admin_city": "광양시",
                    "admin_locality": "중동",
                },
                {
                    "station_id": 168,
                    "station_name": "여수",
                    "station_address": "전라남도 여수시 중앙동",
                    "station_latitude": 34.7393,
                    "station_longitude": 127.7406,
                    "admin_province": "전라남도",
                    "admin_city": "여수시",
                    "admin_locality": "중앙동",
                },
            ]
        )
    )


def test_administrative_region_is_separate_from_weather_station_name():
    area = parse_administrative_area("전남 광양시 도이동 775")
    assert area.province == "전라남도"
    assert area.city == "광양시"
    assert area.locality == "도이동"

    match = _stations().resolve(
        address="전남 광양시 도이동 775",
        latitude=None,
        longitude=None,
    )
    assert match.station_id == 266
    assert match.method == "administrative_area_exact"
    assert not match.review_required


def test_station_catalog_quarantines_an_address_without_a_defensible_match():
    match = _stations().resolve(
        address="전국",
        latitude=None,
        longitude=None,
    )
    assert match.station_id is None
    assert match.review_required
    assert "insufficient" in str(match.reason)


def test_nationwide_builder_does_not_filter_to_legacy_company_or_date(tmp_path: Path):
    partitions = []
    for company, plant, timestamp in (
        ("kospo", "남부A", "2019-01-01"),
        ("koen", "남동B", "2025-01-01"),
    ):
        path = tmp_path / f"{company}.csv.gz"
        pd.DataFrame(
            {
                "timestamp": [timestamp],
                "company": [company],
                "plant_id": [f"{company}:{plant}"],
                "plant": [plant],
                "energy_source": ["solar"],
                "generation_mwh": [1.0],
            }
        ).to_csv(path, index=False, compression="gzip")
        partitions.append(path)

    registry = pd.DataFrame(
        {
            "company": ["kospo", "koen"],
            "plant": ["남부A", "남동B"],
            "energy_source": ["solar", "solar"],
            "capacity_mw": [1.0, 1.0],
            "tilt_deg": [20.0, 20.0],
            "admin_province": ["부산광역시", "전라남도"],
            "admin_city": ["사하구", "광양시"],
            "weather_station_id": [159, 266],
            "weather_station_name": ["부산", "광양시"],
            "weather_mapping_method": ["reviewed_legacy_mapping", "administrative_area_exact"],
            "weather_mapping_confidence": ["high", "high"],
            "weather_mapping_review_required": [False, False],
            "model_ready_status": ["eligible", "eligible"],
        }
    )
    builder = NationwideModelDatasetBuilder(tmp_path, PlantMetadataCatalog([]))
    result = builder._from_standardized_generation(partitions, registry)
    assert set(result["company"]) == {"kospo", "koen"}
    assert result["timestamp"].min() == pd.Timestamp("2019-01-01")
    assert result["timestamp"].max() == pd.Timestamp("2025-01-01")


def test_nationwide_builder_replaces_cumulative_revision_instead_of_summing_it(tmp_path: Path):
    paths = []
    for index, value in enumerate((1.0, 1.2)):
        path = tmp_path / f"revision_{index}.csv.gz"
        pd.DataFrame(
            {
                "timestamp": ["2025-01-01"],
                "company": ["kospo"],
                "plant_id": ["kospo:용수리#1"],
                "plant": ["용수리"],
                "energy_source": ["solar"],
                "generation_mwh": [value],
            }
        ).to_csv(path, index=False, compression="gzip")
        paths.append(path)
    registry = pd.DataFrame(
        {
            "company": ["kospo"],
            "plant": ["용수리"],
            "energy_source": ["solar"],
            "capacity_mw": [2.0],
            "tilt_deg": [20.0],
            "admin_province": ["제주특별자치도"],
            "admin_city": ["제주시"],
            "weather_station_id": [185],
            "weather_station_name": ["고산"],
            "weather_mapping_method": ["reviewed_legacy_mapping"],
            "weather_mapping_confidence": ["high"],
            "weather_mapping_review_required": [False],
            "model_ready_status": ["eligible"],
        }
    )
    builder = NationwideModelDatasetBuilder(tmp_path, PlantMetadataCatalog([]))
    result = builder._from_standardized_generation(paths, registry)
    assert result.iloc[0]["generation_mwh"] == 1.2
    assert builder._reconciliation["revision_conflict_keys"] == 1


def test_nationwide_builder_prefers_explicit_latest_snapshot_date(tmp_path: Path):
    paths = []
    for snapshot, value in (("20260228", 1.2), ("20240101", 1.0)):
        path = tmp_path / f"plant_{snapshot}.csv.gz"
        pd.DataFrame(
            {
                "timestamp": ["2024-01-01"],
                "company": ["kospo"],
                "plant_id": ["kospo:용수리#1"],
                "plant": ["용수리"],
                "energy_source": ["solar"],
                "generation_mwh": [value],
            }
        ).to_csv(path, index=False, compression="gzip")
        paths.append(path)
    registry = pd.DataFrame(
        {
            "company": ["kospo"],
            "plant": ["용수리"],
            "energy_source": ["solar"],
            "capacity_mw": [2.0],
            "tilt_deg": [20.0],
            "admin_province": ["전라남도"],
            "admin_city": ["해남군"],
            "weather_station_id": [261],
            "weather_station_name": ["해남"],
            "weather_mapping_method": ["reviewed_legacy_mapping"],
            "weather_mapping_confidence": ["high"],
            "weather_mapping_review_required": [False],
            "model_ready_status": ["eligible"],
        }
    )
    builder = NationwideModelDatasetBuilder(tmp_path, PlantMetadataCatalog([]))
    result = builder._from_standardized_generation(paths, registry)
    assert result.iloc[0]["generation_mwh"] == 1.2


def test_model_ready_partitions_are_replaced_by_company_and_year(tmp_path: Path):
    frame = pd.DataFrame(
        {
            "timestamp": pd.to_datetime(["2024-12-31", "2025-01-01"]),
            "company": ["koen", "koen"],
            "generation_mwh": [1.0, 2.0],
        }
    )
    destination = tmp_path / "parts"
    first = NationwideModelDatasetBuilder._write_model_partitions(frame, destination)
    second = NationwideModelDatasetBuilder._write_model_partitions(frame.iloc[[1]], destination)
    assert len(first) == 2
    assert len(second) == 1
    assert len(list(destination.rglob("*.csv.gz"))) == 1
