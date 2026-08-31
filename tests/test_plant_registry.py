from pathlib import Path

import pandas as pd
import pytest

from solar_forecast.collectors.metadata import PlantMetadata, PlantMetadataCatalog
from solar_forecast.features.registry import (
    KmaStationCatalog,
    NationwidePlantRegistryBuilder,
    ReviewedStationMapping,
    ReviewedStationMappingCatalog,
    parse_administrative_area,
)
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


def test_reviewed_alias_keeps_one_physical_plant_identity_and_filters_ess_capacity():
    catalog = PlantMetadataCatalog(
        [
            PlantMetadata("iwest", "화순풍력", "wind", 16.0, None, "전남 화순군", None),
            PlantMetadata("iwest", "화순풍력", "storage", 1.25, None, "전남 화순군", None),
        ]
    )
    assert catalog.canonical_plant("iwest", "영암에프원태양광b") == "영암F1 태양광"
    assert catalog.canonical_plant("iwest", "(군산)영암F1태양광") == "영암F1 태양광"
    wind = catalog.lookup("iwest", "화순풍력발전", energy_source="wind", aggregate=True)
    assert wind is not None
    assert wind.capacity_mw == 16.0


def test_reviewed_station_mapping_catalog_is_explicit_and_auditable():
    mapping = ReviewedStationMapping("krc", "율치", 259, "https://example.test", "reviewed")
    catalog = ReviewedStationMappingCatalog([mapping])
    assert catalog.get("krc", "율치") == mapping
    assert catalog.get("krc", "영암1차") is None


@pytest.mark.parametrize(
    ("evidence_url", "rationale"),
    [
        ("", ""),
        (None, "reviewed"),
        ("https://example.test", None),
        ("None", "reviewed"),
    ],
)
def test_reviewed_station_mapping_requires_evidence_and_rationale(
    evidence_url: object,
    rationale: object,
):
    with pytest.raises(ValueError, match="evidence_url and rationale"):
        ReviewedStationMapping(
            "kospo",
            "무근거",
            168,
            evidence_url,  # type: ignore[arg-type]
            rationale,  # type: ignore[arg-type]
        )


def test_legacy_station_seed_is_audit_only_and_cannot_make_a_plant_eligible(
    tmp_path: Path,
):
    partition = tmp_path / "plant.csv.gz"
    pd.DataFrame(
        {
            "timestamp": ["2024-01-01T00:00:00"],
            "company": ["kospo"],
            "plant": ["무근거 태양광"],
            "energy_source": ["solar"],
            "capacity_mw": [1.0],
            "tilt_deg": [None],
            "latitude": [None],
            "longitude": [None],
            "address": [None],
        }
    ).to_csv(partition, index=False, compression="gzip")
    legacy = pd.DataFrame(
        {
            "company": ["kospo"],
            "plant": ["무근거 태양광"],
            "station_id": [168],
        }
    )

    registry = NationwidePlantRegistryBuilder(
        PlantMetadataCatalog([]),
        _stations(),
    ).build([partition], tmp_path / "registry.csv", legacy_mapping=legacy)

    row = registry.iloc[0]
    assert row["model_ready_status"] == "quarantined"
    assert row["weather_mapping_method"] == "unresolved"
    assert row["legacy_weather_station_candidate_id"] == 168
    assert row["legacy_weather_station_candidate_status"] == "audit_only_unreviewed"
    assert "audit candidate" in row["model_ready_reason"]


def test_reviewed_mapping_requires_one_station_record_to_cover_generation_dates():
    stations = _stations().stations.copy()
    stations["station_valid_from"] = pd.to_datetime(["2025-04-25", "1900-01-01"])
    stations["station_valid_to"] = pd.to_datetime([None, "2024-12-31"])
    catalog = KmaStationCatalog(stations)

    outside = catalog.resolve(
        address="전라남도 광양시 중동",
        latitude=None,
        longitude=None,
        reviewed_station_id=266,
        generation_start="2013-01-01",
        generation_end="2025-02-28",
    )
    partial = catalog.resolve(
        address="전라남도 여수시 중앙동",
        latitude=None,
        longitude=None,
        reviewed_station_id=168,
        generation_start="2013-01-01",
        generation_end="2025-02-28",
    )
    continued = stations.loc[stations["station_id"].eq(168)].copy()
    continued["station_valid_from"] = pd.Timestamp("2025-01-01")
    continued["station_valid_to"] = pd.NaT
    covered_stations = pd.concat([stations, continued], ignore_index=True)
    covered = KmaStationCatalog(covered_stations).resolve(
        address="전라남도 여수시 중앙동",
        latitude=None,
        longitude=None,
        reviewed_station_id=168,
        generation_start="2013-01-01",
        generation_end="2025-02-28",
    )

    assert outside.station_id is None
    assert outside.review_required
    assert outside.method == "reviewed_mapping_invalid"
    assert partial.station_id is None
    assert partial.review_required
    assert covered.station_id == 168
    assert covered.method == "reviewed_config_mapping"
    assert not covered.review_required


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
            "weather_mapping_method": ["administrative_area_exact", "administrative_area_exact"],
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


@pytest.mark.parametrize(
    "mapping_method",
    ["reviewed_legacy_mapping", "unknown_future_mapping", None],
)
def test_nationwide_builder_rejects_unapproved_mapping_methods(
    tmp_path: Path,
    mapping_method: object,
):
    partition = tmp_path / "legacy.csv.gz"
    pd.DataFrame(
        {
            "timestamp": ["2024-01-01"],
            "company": ["kospo"],
            "plant_id": ["kospo:구형"],
            "plant": ["구형"],
            "energy_source": ["solar"],
            "generation_mwh": [1.0],
        }
    ).to_csv(partition, index=False, compression="gzip")
    registry = pd.DataFrame(
        {
            "company": ["kospo"],
            "plant": ["구형"],
            "energy_source": ["solar"],
            "weather_station_id": [168],
            "weather_mapping_method": [mapping_method],
            "weather_mapping_review_required": [False],
            "model_ready_status": ["eligible"],
        }
    )

    builder = NationwideModelDatasetBuilder(tmp_path, PlantMetadataCatalog([]))
    with pytest.raises(ValueError, match="legacy station seeds require"):
        builder._from_standardized_generation([partition], registry)


def test_nationwide_builder_rejects_null_reviewed_evidence(tmp_path: Path):
    partition = tmp_path / "reviewed.csv.gz"
    pd.DataFrame(
        {
            "timestamp": ["2024-01-01"],
            "company": ["kospo"],
            "plant_id": ["kospo:검토"],
            "plant": ["검토"],
            "energy_source": ["solar"],
            "generation_mwh": [1.0],
        }
    ).to_csv(partition, index=False, compression="gzip")
    registry = pd.DataFrame(
        {
            "company": ["kospo"],
            "plant": ["검토"],
            "energy_source": ["solar"],
            "weather_station_id": [168],
            "weather_mapping_method": ["reviewed_config_mapping"],
            "weather_mapping_review_required": [False],
            "weather_mapping_evidence_url": [None],
            "weather_mapping_rationale": ["reviewed"],
            "model_ready_status": ["eligible"],
        }
    )

    builder = NationwideModelDatasetBuilder(tmp_path, PlantMetadataCatalog([]))
    with pytest.raises(ValueError, match="explicit reviewed evidence"):
        builder._from_standardized_generation([partition], registry)


def test_official_partitions_do_not_require_a_legacy_mapping_file(tmp_path: Path):
    builder = NationwideModelDatasetBuilder(tmp_path, PlantMetadataCatalog([]))

    assert builder._read_optional_legacy_generation(tmp_path / "missing-val.csv") is None


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
            "weather_mapping_method": ["administrative_area_exact"],
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
            "weather_mapping_method": ["administrative_area_exact"],
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
