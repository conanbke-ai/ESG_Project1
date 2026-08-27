import json

import pandas as pd
import pytest

from solar_forecast.collectors.candidates import (
    CandidateAcceptancePolicy,
    KrcYeongamCandidateIntakeService,
)
from solar_forecast.collectors.normalization import KrcYeongamGenerationNormalizer


def _krc_row(month: int, day: int, plant: str, value: float = 1_000) -> dict[str, object]:
    row: dict[str, object] = {"월": month, "일": day, "시설명": plant}
    row.update({f" {hour}시 ": value for hour in range(1, 25)})
    row["계"] = value * 24
    return row


def test_krc_normalizer_maps_entity_capacity_and_hourly_mwh():
    source = pd.DataFrame([_krc_row(1, 1, "영암1차")])
    result = KrcYeongamGenerationNormalizer().transform(source, year=2025)

    assert len(result) == 24
    assert result.iloc[0]["timestamp"] == pd.Timestamp("2025-01-01 00:00")
    assert result.iloc[-1]["timestamp"] == pd.Timestamp("2025-01-01 23:00")
    assert result["plant_id"].eq("krc:영암1차").all()
    assert result["generation_mwh"].eq(1.0).all()
    assert result["capacity_mw"].astype(float).eq(1.4916).all()


def test_krc_normalizer_quarantines_files_without_entity_column():
    source = pd.DataFrame([{**_krc_row(1, 1, "영암1차"), "시설명": None}]).drop(
        columns="시설명"
    )
    with pytest.raises(ValueError, match="entity-safe"):
        KrcYeongamGenerationNormalizer().transform(source, year=2021)


def test_candidate_intake_separates_generation_gate_from_weather_admission(tmp_path):
    source_dir = tmp_path / "raw"
    output_dir = tmp_path / "standardized"
    source_dir.mkdir()
    rows = [
        _krc_row(1, day, plant)
        for day in range(1, 5)
        for plant in ("영암1차", "영암2차", "율치")
    ]
    pd.DataFrame(rows).to_csv(
        source_dir / "한국농어촌공사_영암 태양광 발전소 발전량 현황_20221231.csv",
        index=False,
        encoding="cp949",
    )
    legacy = pd.DataFrame([{key: value for key, value in _krc_row(1, 1, "영암1차").items() if key != "시설명"}])
    legacy.to_csv(
        source_dir / "한국농어촌공사_영암 태양광 발전소 발전량 현황_20211231.csv",
        index=False,
        encoding="cp949",
    )

    result = KrcYeongamCandidateIntakeService(
        source_dir,
        output_dir,
        CandidateAcceptancePolicy(
            minimum_days=4,
            gap_hours=0,
            maximum_hourly_capacity_factor=1.10,
        ),
    ).run()

    assert result.rows == 4 * 3 * 24
    assert result.plants == 3
    assert result.status == "generation_ready_for_registry"
    assert [item.status for item in result.source_files] == [
        "quarantined",
        "accepted_for_generation_audit",
    ]
    assert all(profile.generation_gate_passed for profile in result.profiles)
    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["unit_review_required"] is False
    assert manifest["unit_validation"]["passed"] is True
    assert manifest["weather_join"]["ready"] is False
    assert set(manifest["split_protocol"]["rows"]) == {
        "train",
        "validation",
        "calibration",
        "test",
    }
