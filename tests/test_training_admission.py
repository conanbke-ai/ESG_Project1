from __future__ import annotations

from solar_forecast.evaluation.training_admission import classify_training_population


def _plant(plant_id: str, *, structural: bool = False, coverage: float = 0.98, run: int = 1000,
           weather: float = 0.97, rows: tuple[int, int, int, int] = (1000, 200, 150, 200)) -> dict:
    splits = {}
    for name, count in zip(("train", "validation", "calibration", "test"), rows):
        splits[name] = {"generation_weather_overlap": {"rows": count}}
    return {
        "plant_id": plant_id,
        "status": "STRUCTURAL_REJECT" if structural else "CANDIDATE_REQUIRES_THRESHOLD_REVIEW",
        "hard_reject_reasons": ["empty_overlap_splits:test"] if structural else [],
        "generation_weather_overlap": {
            "hourly_coverage": coverage,
            "longest_strict_run_hours": run,
        },
        "weather_column_coverage": {
            "temperature_c": {"coverage": weather},
            "solar_irradiance_mj_m2": {"coverage": weather},
        },
        "splits": splits,
    }


def _eligibility(*plants: dict) -> dict:
    return {"contract": "solar-training-data-eligibility.v1", "plants": list(plants)}


def _policy(status: str = "frozen") -> dict:
    return {
        "contract": "solar-training-admission-policy.v1",
        "version": "test",
        "status": status,
        "fixed_start_year": None,
        "admit_thresholds": {
            "min_overlap_hourly_coverage": 0.95 if status == "frozen" else None,
            "min_longest_strict_run_hours": 500 if status == "frozen" else None,
            "min_weather_column_coverage": 0.90 if status == "frozen" else None,
            "min_split_rows": {
                "train": 500 if status == "frozen" else None,
                "validation": 100 if status == "frozen" else None,
                "calibration": 100 if status == "frozen" else None,
                "test": 100 if status == "frozen" else None,
            },
        },
        "reject_thresholds": {
            "min_overlap_hourly_coverage": 0.80 if status == "frozen" else None,
            "min_longest_strict_run_hours": 168 if status == "frozen" else None,
            "min_weather_column_coverage": 0.70 if status == "frozen" else None,
            "min_split_rows": {
                "train": 200 if status == "frozen" else None,
                "validation": 24 if status == "frozen" else None,
                "calibration": 24 if status == "frozen" else None,
                "test": 24 if status == "frozen" else None,
            },
        },
    }


def test_draft_policy_holds_structurally_valid_candidates() -> None:
    report = classify_training_population(
        _eligibility(_plant("good"), _plant("bad", structural=True)),
        _policy("draft"),
    )
    states = {item["plant_id"]: item["status"] for item in report["decisions"]}
    assert states == {"good": "HOLD", "bad": "REJECT"}
    assert report["final_training_selection_ready"] is False
    assert "admission_policy_not_frozen" in report["selection_blockers"]


def test_frozen_policy_separates_admit_hold_and_reject() -> None:
    report = classify_training_population(
        _eligibility(
            _plant("admit"),
            _plant("hold", coverage=0.90),
            _plant("reject", coverage=0.70),
        ),
        _policy("frozen"),
    )
    states = {item["plant_id"]: item["status"] for item in report["decisions"]}
    assert states == {"admit": "ADMITTED", "hold": "HOLD", "reject": "REJECT"}
    assert report["final_training_selection_ready"] is False
    assert report["counts"] == {"total": 3, "admitted": 1, "hold": 1, "rejected": 1}


def test_frozen_population_is_ready_only_without_holds() -> None:
    report = classify_training_population(
        _eligibility(_plant("admit"), _plant("reject", structural=True)),
        _policy("frozen"),
    )
    assert report["admitted_plant_ids"] == ["admit"]
    assert report["rejected_plant_ids"] == ["reject"]
    assert report["hold_plant_ids"] == []
    assert report["final_training_selection_ready"] is True
    assert report["selection_blockers"] == []
