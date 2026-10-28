from __future__ import annotations

from pathlib import Path

import pandas as pd

from solar_forecast.datasets.repository import DatasetLoadPolicy, DatasetRepository


def test_training_loader_applies_admitted_plant_allow_list_without_mutating_source(tmp_path: Path) -> None:
    source = tmp_path / "gold.csv"
    original = pd.DataFrame(
        [
            {
                "plant_id": "plant-a",
                "energy_source": "solar",
                "quality_train_eligible": True,
                "timestamp": "2024-01-01 00:00:00",
                "generation_mwh": 1.0,
            },
            {
                "plant_id": "plant-b",
                "energy_source": "solar",
                "quality_train_eligible": True,
                "timestamp": "2024-01-01 00:00:00",
                "generation_mwh": 2.0,
            },
            {
                "plant_id": "plant-a",
                "energy_source": "wind",
                "quality_train_eligible": True,
                "timestamp": "2024-01-01 01:00:00",
                "generation_mwh": 3.0,
            },
        ]
    )
    original.to_csv(source, index=False)
    before = source.read_bytes()

    _, frame, report = DatasetRepository(tmp_path).load_training_frame(
        source,
        columns=["plant_id", "timestamp", "generation_mwh"],
        numeric_columns=["generation_mwh"],
        equals_filters={"energy_source": "solar"},
        allowed_values_filters={"plant_id": ["plant-a"]},
        truthy_filter="quality_train_eligible",
        policy=DatasetLoadPolicy(chunk_rows=2, memory_limit_mb=64, numeric_dtype="float64"),
    )

    assert frame["plant_id"].astype(str).tolist() == ["plant-a"]
    assert frame["generation_mwh"].tolist() == [1.0]
    assert report.scanned_rows == 3
    assert report.retained_rows == 1
    assert source.read_bytes() == before


def test_training_loader_rejects_empty_allow_list(tmp_path: Path) -> None:
    source = tmp_path / "gold.csv"
    pd.DataFrame(
        [
            {
                "plant_id": "plant-a",
                "energy_source": "solar",
                "quality_train_eligible": True,
                "timestamp": "2024-01-01 00:00:00",
                "generation_mwh": 1.0,
            }
        ]
    ).to_csv(source, index=False)

    try:
        DatasetRepository(tmp_path).load_training_frame(
            source,
            columns=["plant_id", "timestamp", "generation_mwh"],
            numeric_columns=["generation_mwh"],
            allowed_values_filters={"plant_id": []},
        )
    except ValueError as exc:
        assert "empty allow-list" in str(exc)
    else:
        raise AssertionError("Empty admitted allow-list must fail closed")
