from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path

import pytest

from solar_forecast.reporting.national_inventory import (
    InventoryContentError,
    InventoryIntegrityError,
    InventorySchemaError,
    NationalInventoryService,
    REQUIRED_COLUMNS,
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
        "canonical_regions": 17,
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
