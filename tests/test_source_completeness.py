from datetime import date
import json

import pandas as pd

from solar_forecast.collectors.generation_normalizers import EwpTrainingNormalizer
from solar_forecast.datasets.source_completeness import audit_source_completeness
from solar_forecast.infrastructure.artifact_store import sha256_file


def test_koen_source_audit_keeps_gaps_conflicts_and_replay_provenance(tmp_path):
    archive = tmp_path / "archive"
    raw = tmp_path / "raw"
    retained = archive / "한국남동발전" / "2024" / "may.csv"
    retained.parent.mkdir(parents=True)
    replay = raw / "koen" / "replay.csv"
    replay.parent.mkdir(parents=True)
    revisions = raw / "koen" / "revision.csv"
    base = {"발전구분": "발전소", "호기": 1, **{f"{hour}시 발전량(kWh)": 0 for hour in range(1, 25)}}
    original = [{**base, "일자": day} for day in ("2024-05-01", "2024-05-03", "2024-05-03")]
    pd.DataFrame(original).to_csv(retained, index=False, encoding="cp949")
    replay.write_bytes(retained.read_bytes())
    pd.DataFrame([{**original[-1], "12시 발전량(kWh)": 10}]).to_csv(revisions, index=False)
    index = raw / "runs" / "run-one" / "collection_manifest.json"
    index.parent.mkdir(parents=True)
    index.write_text(json.dumps({
        "results": [{"source": "koen", "status": "partial", "rows": 24, "windows": [
            {"start_date": "2024-05-03", "end_date": "2024-05-03", "status": "downloaded",
             "source_rows": 1, "path": str(revisions)},
            {"start_date": "2024-06-01", "end_date": "2024-06-01", "status": "empty_csv",
             "source_rows": 0},
        ]}],
        "file_artifacts": [{"path": str(revisions), "sha256": sha256_file(revisions),
                            "role": "provider_original_bronze"}],
    }), encoding="utf-8")

    report = audit_source_completeness(archive, raw, date(2026, 11, 25))

    koen = report["providers"]["koen"]
    by_file = {item["path"]: item for item in koen["files"]}
    assert by_file[str(retained)]["repeated_source_key_rows"] == 1
    assert by_file[str(replay)]["duplicate_of"] == str(retained)
    assert by_file[str(revisions)]["indexed_by_collection_manifest"]
    assert report["collection_indexes"][0]["originals"][0]["hash_present_in_scanned_inputs"]
    index_result = report["collection_indexes"][0]["results"][0]
    assert index_result["status"] == "partial"
    assert [window["status"] for window in index_result["windows"]] == ["downloaded", "empty_csv"]
    coverage = koen["coverage_by_layout"]["plant_meter_hour"]["2024"]
    assert coverage["unique_hour_keys"] == 48
    assert coverage["same_value_repeat_rows"] == 47
    assert coverage["conflicting_value_hour_keys"] == 1
    assert coverage["conflict_sample"][0]["source_files"] == sorted([str(retained), str(revisions)])
    assert coverage["entities"][0]["gaps"] == [
        {"first": "2024-05-02T00:00:00", "last": "2024-05-02T23:00:00", "hours": 24},
    ]
    assert retained.read_bytes() == replay.read_bytes()


def test_kospo_same_bytes_with_distinct_filename_plant_are_distinct_entities(tmp_path):
    archive = tmp_path / "archive"
    kospo = archive / "한국남부발전"
    kospo.mkdir(parents=True)
    data = pd.DataFrame([{"년월일": "2026-02-02", "호기": 1, **{str(hour): 0 for hour in range(1, 25)}}])
    for name in ("신풍리", "용수리"):
        data.to_csv(kospo / f"[{name}]태양광.csv", index=False)
    index = tmp_path / "unknown_index.json"
    index.write_text('{"acquisition": []}', encoding="utf-8")

    report = audit_source_completeness(archive, tmp_path / "raw", date(2026, 11, 25), collection_indexes=(index,))

    kospo_2026 = report["providers"]["kospo"]["coverage_by_layout"]["plant_meter_hour"]["2026"]
    assert kospo_2026["unique_hour_keys"] == 48
    assert {item["entity"] for item in kospo_2026["entities"]} == {"신풍리#1", "용수리#1"}
    assert report["collection_indexes"][0]["error"].startswith("Unrecognized collection index schema")


def test_ewp_nationwide_revisions_stay_at_region_scope(tmp_path):
    ewp = tmp_path / "raw" / "ewp"
    ewp.mkdir(parents=True)
    base = dict.fromkeys(EwpTrainingNormalizer.column_map, 0)
    base.update({"시도명": "경상남도", "발전일자": "2024-09-05 13:00", "발전량(MWh)": 5})
    pd.DataFrame([base, base, {**base, "발전량(MWh)": 5.5}]).to_csv(ewp / "training.csv", index=False)

    report = audit_source_completeness(tmp_path / "archive", ewp.parent, date(2026, 11, 25))

    coverage = report["providers"]["ewp"]["coverage_by_layout"]
    assert set(coverage) == {"region_hour"}
    assert coverage["region_hour"]["2024"]["unique_hour_keys"] == 1
    assert coverage["region_hour"]["2024"]["same_value_repeat_rows"] == 1
    assert coverage["region_hour"]["2024"]["conflicting_value_hour_keys"] == 1
