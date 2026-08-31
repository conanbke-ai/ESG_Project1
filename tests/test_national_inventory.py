from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path

import pytest

from solar_forecast.reporting.national_inventory import (
    AdministrativeRegionReference,
    InventoryConfigurationError,
    InventoryContentError,
    InventoryIntegrityError,
    InventorySchemaError,
    NationalInventoryService,
    REQUIRED_COLUMNS,
    build_national_inventory,
    canonical_location,
    canonical_region,
    source_region_conflict,
)


def _write_csv(path: Path, rows: list[list[str]], *, encoding: str = "cp949") -> str:
    import csv

    with path.open("w", encoding=encoding, newline="") as target:
        writer = csv.writer(target)
        writer.writerow(REQUIRED_COLUMNS)
        writer.writerows(rows)
    return sha256(path.read_bytes()).hexdigest()


def _config(path: Path, digest: str, *, encoding: str = "cp949") -> dict[str, object]:
    return {
        "dataset_id": "test-kpx-solar",
        "provider": "한국전력거래소",
        "source_system": "EPSIS",
        "source_url": "https://example.test/epsis",
        "local_path": path.name,
        "reference_date": "2026-08-05",
        "downloaded_at": "2026-08-28T10:00:00+09:00",
        "expected_sha256": digest,
        "encoding": encoding,
        "capacity_unit": "MW",
        "record_unit": "generator_registration_row",
        "scope": "전국 태양광 발전기 등록 레코드",
        "limitations": ["한 행은 물리적 발전소 개소와 다를 수 있다."],
    }


def _write_location_overrides(
    path: Path,
    digest: str,
    overrides: list[dict[str, object]],
) -> None:
    path.write_text(
        json.dumps(
            {
                "schema_version": 1,
                "dataset_id": "test-kpx-solar",
                "source_sha256": digest,
                "reviewed_at": "2026-08-31",
                "unresolved_policy": (
                    "keep_source_region_and_quarantine_subregion"
                ),
                "overrides": overrides,
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def _approved_override(**changes: object) -> dict[str, object]:
    values: dict[str, object] = {
        "id": "move-suncheon",
        "source_row_numbers": [2],
        "old_region": "전북",
        "old_subregion": "전라남도 순천시",
        "new_region": "전남광주통합특별시",
        "new_subregion": "전남광주통합특별시 순천시",
        "expected_records": 1,
        "expected_capacity_mw": "1.5",
        "evidence_provider": "공공데이터포털",
        "evidence_url": "https://www.data.go.kr/data/15107742/standard.do",
        "evidence_reference_date": "2026-08-13",
        "match_method": "exact_name_capacity_address",
        "confidence": "high",
        "review_status": "approved",
        "reason": "공식 허가 데이터의 설비명·용량·주소가 일치함",
    }
    values.update(changes)
    return values


def test_cp949_inventory_keeps_duplicates_and_validates_footer_and_coordinates(tmp_path):
    source = tmp_path / "inventory.csv"
    row_a = [
        "사업자A",
        "발전기A",
        "1",
        "1.5",
        "비회원",
        "비중앙",
        "태양에너지",
        "신재생",
        "발전사업",
        "전북",
        "전북특별자치도 익산시",
    ]
    rows = [
        row_a,
        row_a.copy(),
        [
            "사업자B",
            "？발전기B",
            "1",
            "0",
            "비회원",
            "비중앙",
            "태양에너지",
            "신재생",
            "발전사업",
            "강원도",
            "강원도  춘천시",
        ],
        [
            "사업자C",
            "발전기C",
            "1",
            "-0.1",
            "비회원",
            "비중앙",
            "태양에너지",
            "신재생",
            "발전사업",
            "서울",
            "",
        ],
        ["", "통합", "4", "2.9", "", "", "", "", "", "", ""],
    ]
    digest = _write_csv(source, rows)
    cache = {
        "전북특별자치도 익산시": [35.9487, 126.9579],
        "강원특별자치도 춘천시": [37.8813, 127.7298],
    }

    payload = NationalInventoryService.from_config(
        _config(source, digest),
        project_root=tmp_path,
        coordinate_cache=cache,
    ).build()
    inventory = payload["national_inventory"]

    assert inventory["summary"] == {
        "generator_records": 4,
        "total_capacity_mw": 2.9,
        "canonical_regions": 16,
        "regions_with_records": 3,
        "source_subregion_labels": 2,
        "subregions": 3,
        "located_subregions": 3,
    }
    assert inventory["source"]["sha256_verified"] is True
    assert inventory["source"]["sha256"] == digest
    assert inventory["quality"]["exact_duplicate_records"] == 1
    assert inventory["quality"]["exact_duplicate_capacity_mw"] == 1.5
    assert inventory["quality"]["duplicates_retained"] is True
    assert inventory["quality"]["zero_capacity_records"] == 1
    assert inventory["quality"]["negative_capacity_records"] == 1
    assert inventory["quality"]["fullwidth_question_mark_cells"] == 1
    assert inventory["quality"]["coordinate_basis_counts"] == {
        "exact": 1,
        "normalized": 1,
        "province_centroid": 1,
    }
    assert inventory["quality"]["footer"] == {
        "found": True,
        "excluded_records": 1,
        "declared_record_count": 4,
        "declared_capacity_mw": 2.9,
        "record_count_matches": True,
        "capacity_matches": True,
        "capacity_difference_mw": 0.0,
    }

    regions = {item["region"]: item for item in inventory["regions"]}
    assert regions["전북특별자치도"]["generator_records"] == 2
    assert regions["강원특별자치도"]["generator_records"] == 1
    assert regions["서울특별시"]["generator_records"] == 1
    locations = {item["subregion"]: item for item in inventory["locations"]}
    assert locations["전북특별자치도 익산시"]["generator_records"] == 2
    assert locations["전북특별자치도 익산시"]["coordinate_basis"] == "exact"
    assert locations["강원특별자치도 춘천시"]["coordinate_basis"] == "normalized"
    assert locations["서울특별시"]["coordinate_basis"] == "province_centroid"
    assert not any(item["source_region_conflict"] for item in inventory["locations"])
    json.dumps(payload, ensure_ascii=False, allow_nan=False)


def test_source_region_conflict_is_flagged_without_silent_reassignment():
    assert source_region_conflict("강원특별자치도 강릉시", "경기도") is True
    assert source_region_conflict("전남 영암군", "전라남도") is False
    assert source_region_conflict("영암군", "전라남도") is False


def test_reviewed_location_override_is_applied_before_aggregation(tmp_path):
    source = tmp_path / "inventory.csv"
    rows = [
        [
            "순천사업자",
            "순천태양광",
            "1",
            "1.5",
            "비회원",
            "비중앙",
            "태양에너지",
            "신재생",
            "발전사업",
            "전북",
            "전라남도 순천시",
        ],
        [
            "미확정사업자",
            "동명이설비",
            "1",
            "2.0",
            "비회원",
            "비중앙",
            "태양에너지",
            "신재생",
            "발전사업",
            "전남",
            "전라북도 군산시",
        ],
        [
            "여수사업자",
            "여수태양광",
            "1",
            "0.5",
            "비회원",
            "비중앙",
            "태양에너지",
            "신재생",
            "발전사업",
            "전남",
            "전라남도 여수시",
        ],
        ["", "통합", "3", "4.0", "", "", "", "", "", "", ""],
    ]
    digest = _write_csv(source, rows)
    override_path = tmp_path / "overrides.json"
    _write_location_overrides(
        override_path, digest, [_approved_override()]
    )
    config = _config(source, digest)
    config["location_override_path"] = override_path.name

    inventory = NationalInventoryService.from_config(
        config,
        project_root=tmp_path,
        coordinate_cache={
            "전남광주통합특별시 순천시": [34.9507, 127.4872],
            "전북특별자치도 군산시": [35.9677, 126.7366],
        },
    ).build()["national_inventory"]

    assert inventory["summary"]["generator_records"] == 3
    assert inventory["summary"]["total_capacity_mw"] == 4.0
    regions = {row["region"]: row for row in inventory["regions"]}
    assert regions["전북특별자치도"]["generator_records"] == 0
    assert regions["전남광주통합특별시"]["generator_records"] == 3

    locations = {
        (row["region"], row["subregion"]): row
        for row in inventory["locations"]
    }
    corrected = locations[
        ("전남광주통합특별시", "전남광주통합특별시 순천시")
    ]
    assert corrected["coordinate_basis"] == "exact"
    unresolved = locations[
        ("전남광주통합특별시", "전남광주통합특별시 미확정 지역")
    ]
    assert unresolved["source_region_conflict"] is True
    assert unresolved["coordinate_basis"] == "province_centroid"
    assert not any(
        row["region"] == "전남광주통합특별시"
        and row["subregion"].startswith("전북특별자치도")
        for row in inventory["locations"]
    )

    quality = inventory["quality"]
    assert quality["source_region_conflict_records"] == 2
    assert quality["source_region_conflict_capacity_mw"] == 3.5
    assert quality["reviewed_location_override_records"] == 1
    assert quality["reviewed_location_override_capacity_mw"] == 1.5
    assert quality["unresolved_region_conflict_records"] == 1
    assert quality["unresolved_region_conflict_capacity_mw"] == 2.0
    assert quality["location_overrides"] == {
        "configured": True,
        "verified": True,
        "local_path": "overrides.json",
        "entries": 1,
        "applied_records": 1,
        "applied_capacity_mw": 1.5,
    }


@pytest.mark.parametrize(
    ("override_changes", "message"),
    [
        ({"old_region": "경기"}, "source matcher failed"),
        ({"expected_capacity_mw": "1.6"}, "does not match its contract"),
    ],
)
def test_reviewed_location_override_fails_closed_on_contract_drift(
    tmp_path, override_changes, message
):
    source = tmp_path / "inventory.csv"
    rows = [
        [
            "순천사업자",
            "순천태양광",
            "1",
            "1.5",
            "비회원",
            "비중앙",
            "태양에너지",
            "신재생",
            "발전사업",
            "전북",
            "전라남도 순천시",
        ],
        ["", "통합", "1", "1.5", "", "", "", "", "", "", ""],
    ]
    digest = _write_csv(source, rows)
    override_path = tmp_path / "overrides.json"
    _write_location_overrides(
        override_path,
        digest,
        [_approved_override(**override_changes)],
    )
    config = _config(source, digest)
    config["location_override_path"] = override_path.name

    with pytest.raises(InventoryContentError, match=message):
        NationalInventoryService.from_config(
            config, project_root=tmp_path, coordinate_cache={}
        ).build()


def test_reviewed_location_override_rejects_a_different_source_hash(tmp_path):
    source = tmp_path / "inventory.csv"
    rows = [
        [
            "순천사업자",
            "순천태양광",
            "1",
            "1.5",
            "비회원",
            "비중앙",
            "태양에너지",
            "신재생",
            "발전사업",
            "전북",
            "전라남도 순천시",
        ],
        ["", "통합", "1", "1.5", "", "", "", "", "", "", ""],
    ]
    digest = _write_csv(source, rows)
    override_path = tmp_path / "overrides.json"
    _write_location_overrides(
        override_path, "0" * 64, [_approved_override()]
    )
    config = _config(source, digest)
    config["location_override_path"] = override_path.name

    with pytest.raises(InventoryConfigurationError, match="source_sha256"):
        NationalInventoryService.from_config(
            config, project_root=tmp_path, coordinate_cache={}
        ).build()


def test_production_location_review_contract_preserves_totals_and_nesting():
    root = Path(__file__).resolve().parents[1]
    inventory = build_national_inventory(root)["national_inventory"]

    assert inventory["summary"]["generator_records"] == 188_594
    assert inventory["summary"]["total_capacity_mw"] == pytest.approx(
        33_059.516180
    )
    assert inventory["summary"]["subregions"] == 258
    quality = inventory["quality"]
    assert quality["source_region_conflict_records"] == 51
    assert quality["source_region_conflict_capacity_mw"] == pytest.approx(
        53.289725
    )
    assert quality["reviewed_location_override_records"] == 47
    assert quality["reviewed_location_override_capacity_mw"] == pytest.approx(
        48.895470
    )
    assert quality["unresolved_region_conflict_records"] == 4
    assert quality["unresolved_region_conflict_capacity_mw"] == pytest.approx(
        4.394255
    )
    assert quality["location_overrides"]["verified"] is True
    assert quality["location_overrides"]["entries"] == 13

    assert not any(
        source_region_conflict(row["subregion"], row["region"])
        for row in inventory["locations"]
    )
    unresolved = {
        (row["region"], row["subregion"], row["generator_records"])
        for row in inventory["locations"]
        if row["source_region_conflict"]
    }
    assert unresolved == {
        ("경상남도", "경상남도 미확정 지역", 1),
        ("경상북도", "경상북도 미확정 지역", 1),
        ("서울특별시", "서울특별시 미확정 지역", 1),
        (
            "전남광주통합특별시",
            "전남광주통합특별시 미확정 지역",
            1,
        ),
    }

    locations = {
        (row["region"], row["subregion"]): row
        for row in inventory["locations"]
    }
    for key in (
        ("대구광역시", "대구광역시 군위군"),
        ("서울특별시", "서울특별시 성동구"),
        ("전북특별자치도", "전북특별자치도 군산시"),
        ("충청남도", "충청남도 금산군"),
        ("경상북도", "경상북도 의성군"),
        ("전남광주통합특별시", "전남광주통합특별시 고흥군"),
    ):
        assert key in locations
        assert locations[key]["source_region_conflict"] is False


def test_effective_dated_reference_merges_legacy_gwangju_jeonnam_names():
    root = Path(__file__).resolve().parents[1]
    reference = AdministrativeRegionReference.from_json(
        root / "config/administrative_regions_20260701.json"
    )

    assert len(reference.canonical_regions) == 16
    assert reference.effective_date == "2026-07-01"
    assert reference.aliases["광주광역시"] == "전남광주통합특별시"
    assert reference.aliases["전라남도"] == "전남광주통합특별시"
    assert (
        reference.location_search_aliases["백령도"]
        == "인천광역시 옹진군"
    )
    assert (
        reference.location_search_aliases["울릉도"]
        == "경상북도 울릉군"
    )
    assert canonical_region("전남") == "전남광주통합특별시"
    assert canonical_location(
        "전라남도 영암군", "전남광주통합특별시"
    ) == "전남광주통합특별시 영암군"


def test_natural_place_aliases_do_not_reassign_generator_name_substrings(tmp_path):
    source = tmp_path / "islands.csv"
    rows = [
        [
            "연천사업자",
            "연천백령1호 태양광발전소",
            "1",
            "0.5",
            "비회원",
            "비중앙",
            "태양에너지",
            "신재생",
            "발전사업",
            "경기",
            "경기도 연천군",
        ],
        [
            "옹진사업자",
            "영흥 태양광발전소",
            "1",
            "1.0",
            "비회원",
            "비중앙",
            "태양에너지",
            "신재생",
            "발전사업",
            "인천",
            "인천광역시 옹진군",
        ],
        ["", "통합", "2", "1.5", "", "", "", "", "", "", ""],
    ]
    digest = _write_csv(source, rows)

    inventory = NationalInventoryService.from_config(
        _config(source, digest), project_root=tmp_path, coordinate_cache={}
    ).build()["national_inventory"]
    locations = {
        (row["region"], row["subregion"]): row
        for row in inventory["locations"]
    }

    assert locations[("경기도", "경기도 연천군")]["generator_records"] == 1
    assert (
        locations[("인천광역시", "인천광역시 옹진군")][
            "generator_records"
        ]
        == 1
    )
    assert inventory["location_search_aliases"]["백령도"] == "인천광역시 옹진군"
    assert not any("울릉군" in row["subregion"] for row in inventory["locations"])


def test_utf8_replacement_character_metric_and_sha256_failure(tmp_path):
    source = tmp_path / "utf8.csv"
    rows = [
        [
            "사업자�",
            "발전기",
            "1",
            "1",
            "비회원",
            "비중앙",
            "태양에너지",
            "신재생",
            "발전사업",
            "제주",
            "제주특별자치도 제주시",
        ],
        ["", "통합", "1", "1", "", "", "", "", "", "", ""],
    ]
    digest = _write_csv(source, rows, encoding="utf-8-sig")
    config = _config(source, digest, encoding="utf-8-sig")
    payload = NationalInventoryService.from_config(
        config, project_root=tmp_path, coordinate_cache={}
    ).build()

    quality = payload["national_inventory"]["quality"]
    assert quality["replacement_character_cells"] == 1
    assert quality["garbled_records"] == 1

    config["expected_sha256"] = "0" * 64
    with pytest.raises(InventoryIntegrityError, match="SHA-256 mismatch"):
        NationalInventoryService.from_config(
            config, project_root=tmp_path, coordinate_cache={}
        ).build()


def test_missing_required_column_is_rejected(tmp_path):
    import csv

    source = tmp_path / "bad-schema.csv"
    with source.open("w", encoding="utf-8", newline="") as target:
        writer = csv.writer(target)
        writer.writerow(REQUIRED_COLUMNS[:-1])
        writer.writerow([""] * (len(REQUIRED_COLUMNS) - 1))
    digest = sha256(source.read_bytes()).hexdigest()

    with pytest.raises(InventorySchemaError, match="세부지역"):
        NationalInventoryService.from_config(
            _config(source, digest, encoding="utf-8"),
            project_root=tmp_path,
            coordinate_cache={},
        ).build()


def test_non_solar_row_is_rejected_instead_of_polluting_national_totals(tmp_path):
    source = tmp_path / "mixed-energy.csv"
    rows = [
        [
            "사업자",
            "풍력발전기",
            "1",
            "3.5",
            "비회원",
            "비중앙",
            "풍력",
            "신재생",
            "발전사업",
            "전남",
            "전라남도 영암군",
        ],
        ["", "통합", "1", "3.5", "", "", "", "", "", "", ""],
    ]
    digest = _write_csv(source, rows)

    with pytest.raises(InventoryContentError, match="CSV row 2.*풍력"):
        NationalInventoryService.from_config(
            _config(source, digest),
            project_root=tmp_path,
            coordinate_cache={},
        ).build()
