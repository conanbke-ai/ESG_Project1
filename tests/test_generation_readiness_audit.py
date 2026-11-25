import hashlib
import json

import pandas as pd
import pytest

from tools.audit_generation_readiness import audit_generation_readiness


def _inputs(tmp_path):
    path = tmp_path / "source.csv"
    pd.DataFrame({
        "timestamp": ["2025-12-31 23:00", "2026-01-01 00:00", "2026-01-01 01:00", "2026-01-01 02:00"],
        "company": ["kospo"] * 4,
        "plant": ["plant"] * 4,
        "plant_id": ["kospo:plant#KPX"] * 4,
        "unit": ["generator-1"] * 4,
        "capacity_mw": [1.0, 1.1, 1.2, None],
    }).to_csv(path, index=False)
    return ((None, path),)


def _review(tmp_path):
    evidence = tmp_path / "reviewed-meter-register.txt"
    evidence.write_text("Retained reviewed physical meter and dated capacities.", encoding="utf-8")
    return {
        "contract": "solar-generation-asset-evidence.v1",
        "evidence": {
            "register": {"path": evidence.name, "sha256": hashlib.sha256(evidence.read_bytes()).hexdigest()}
        },
        "meter_bindings": [{
            "source_series_id": "kospo:plant#KPX",
            "start_kst": "2025-12-31 23:00",
            "end_exclusive_kst": "2026-01-01 02:00",
            "meter_point_id": "meter-a",
            "evidence_id": "register",
        }],
        "capacity_periods": [{
            "source_series_id": "kospo:plant#KPX",
            "start_kst": "2025-12-31 23:00",
            "end_exclusive_kst": "2026-01-01 00:00",
            "capacity_mw": 1.0,
            "evidence_id": "register",
        }, {
            "source_series_id": "kospo:plant#KPX",
            "start_kst": "2026-01-01 00:00",
            "end_exclusive_kst": "2026-01-01 03:00",
            "capacity_mw": 1.1,
            "evidence_id": "register",
        }],
    }


def test_reports_period_conflicts_without_backfilling_capacity(tmp_path):
    inputs = _inputs(tmp_path)
    source = inputs[0][1]
    before = source.read_bytes()
    review = tmp_path / "review.json"
    review.write_text(json.dumps(_review(tmp_path)), encoding="utf-8")

    result = audit_generation_readiness(inputs, review)
    series, = result["source_series"]

    assert series["meter_point_evidence"]["covered_hours"] == 3
    assert series["meter_point_evidence"]["uncovered_hours"] == 1
    assert series["capacity_period_evidence"]["covered_hours"] == 4
    assert series["capacity_period_evidence"]["reported_capacity_conflict_hours"] == 1
    years = series["observed_capacity_by_source"][0]["years"]
    assert [period["year"] for period in years] == [2025, 2026]
    assert years[1]["missing_capacity_rows"] == 1
    assert [value["capacity_mw"] for value in years[1]["observed_values"]] == [1.1, 1.2]
    assert source.read_bytes() == before


def test_reports_source_labels_without_claiming_physical_meter(tmp_path):
    result = audit_generation_readiness(_inputs(tmp_path))
    series, = result["source_series"]
    assert series["generator_or_unit_labels"] == ["generator-1"]
    assert series["meter_point_evidence"]["covered_hours"] == 0
    assert series["meter_point_evidence"]["uncovered_hours"] == 4


@pytest.mark.parametrize("issue", ["bad_hash", "overlapping_capacity", "off_hour"])
def test_rejects_invalid_review(tmp_path, issue):
    data = _review(tmp_path)
    if issue == "bad_hash":
        data["evidence"]["register"]["sha256"] = "0" * 64
        message = "SHA-256 mismatch"
    elif issue == "overlapping_capacity":
        data["capacity_periods"][1]["start_kst"] = "2025-12-31 23:00"
        message = "Overlapping capacity_periods intervals"
    else:
        data["meter_bindings"][0]["end_exclusive_kst"] = "2026-01-01 02:30"
        message = "local KST hour labels"
    review = tmp_path / "review.json"
    review.write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(ValueError, match=message):
        audit_generation_readiness(_inputs(tmp_path), review)


def test_reports_same_physical_meter_binding_across_source_series(tmp_path):
    inputs = _inputs(tmp_path)
    data = _review(tmp_path)
    data["meter_bindings"].append({
        **data["meter_bindings"][0],
        "source_series_id": "kospo:plant#renamed",
        "start_kst": "2026-01-01 00:00",
    })
    review = tmp_path / "review.json"
    review.write_text(json.dumps(data), encoding="utf-8")

    result = audit_generation_readiness(inputs, review)

    assert result["review_series_not_seen"] == ["kospo:plant#renamed"]
    assert result["overlapping_meter_bindings"] == [{
        "meter_point_id": "meter-a",
        "source_series_ids": ["kospo:plant#KPX", "kospo:plant#renamed"],
        "start_kst": "2026-01-01T00:00:00",
        "end_exclusive_kst": "2026-01-01T02:00:00",
    }]
