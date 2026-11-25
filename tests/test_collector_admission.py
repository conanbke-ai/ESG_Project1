from __future__ import annotations

import json

import pandas as pd

from solar_forecast.datasets.collector_admission import CollectedGenerationAdmissionService
from solar_forecast.collectors.generation_normalizers import GENERATION_COLUMNS


def test_collected_generation_admission_accepts_only_plant_hour_contract(tmp_path):
    source_dir = tmp_path / "downloads"
    source_dir.mkdir()
    accepted = source_dir / "accepted.csv"
    rejected = source_dir / "regional_training.csv"

    pd.DataFrame(
        [
            {
                "timestamp": "2025-01-01 00:00:00",
                "company": "koen",
                "plant_id": "koen:test#1",
                "plant": "test",
                "unit": "1",
                "energy_source": "solar",
                "generation_mwh": 0.0,
                "capacity_mw": 1.0,
                "tilt_deg": None,
                "latitude": None,
                "longitude": None,
                "address": None,
                "source_file": "accepted.csv",
            }
        ],
        columns=GENERATION_COLUMNS,
    ).to_csv(accepted, index=False, encoding="utf-8-sig")
    pd.DataFrame(
        [
            {
                "timestamp": "2025-01-01 00:00:00",
                "company": "ewp",
                "region": "전남",
                "generation_mwh": 1.2,
            }
        ]
    ).to_csv(rejected, index=False, encoding="utf-8-sig")

    result = CollectedGenerationAdmissionService(
        source_dir,
        tmp_path / "collector_admission_manifest.json",
    ).run()

    assert result.accepted_paths == (accepted,)
    assert result.accepted_count == 1
    assert result.rejected_count == 1
    assert result.accepted_rows == 1
    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["summary"] == {
        "files": 2,
        "accepted_files": 1,
        "rejected_files": 1,
        "accepted_rows": 1,
    }
    rejection = [
        item for item in manifest["files"] if item["status"] == "rejected"
    ][0]
    assert rejection["reason"].startswith("missing_generation_contract_columns")
