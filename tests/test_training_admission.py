from __future__ import annotations

import pytest

from solar_forecast.evaluation.training_admission import classify_training_population


SPLITS = ("train", "validation", "calibration", "test")


def _plant(
    plant_id: str,
    *,
    structural: bool = False,
    reasons: list[str] | None = None,
    coverage: float = 0.98,
    run: int = 1000,
    weather: float = 0.97,
    rows: tuple[int, int, int, int] = (1000, 200, 150, 200),
    gaps: tuple[int, int, int, int] = (0, 0, 0, 0),
) -> dict:
    splits = {}
    for name, count, gap_count in zip(SPLITS, rows, gaps):
        splits[name] = {
            "generation_weather_overlap": {
                "rows": count,
                "continuity": {"gap_count_gt_1h": gap_count},
            }
        }
    hard_reasons = reasons or (["empty_overlap_splits:test"] if structural else [])
    return {
        "plant_id": plant_id,
        "status": "STRUCTURAL_REJECT" if structural else "CANDIDATE_REQUIRES_THRESHOLD_REVIEW",
        "hard_reject_reasons": hard_reasons,
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


def _eligibility(*plants: dict, contract: str = "solar-training-data-eligibility.v2") -> dict:
    return {
        "contract": contract,
        "historical_rule_recovery": {
            "fixed_minimum_rows_per_split_found": False,
            "historical_3way_extended_to_current_4way": True,
        },
        "plants": list(plants),
    }


def _policy() -> dict:
    return {
        "contract": "solar-training-admission-policy.v2",
        "version": "test-v2",
        "status": "frozen",
        "fixed_start_year": None,
        "current_4way_extension": {
            "require_all_splits_present": True,
            "reject_split_timestamp_gap_gt_hours": 1,
            "require_validation_rows_lte_train_rows": True,
            "require_calibration_rows_lte_train_rows": True,
            "require_test_rows_lte_train_rows": True,
            "fixed_minimum_rows_per_split": None,
        },
    }


def test_v2_admits_only_structurally_valid_candidates() -> None:
    report = classify_training_population(
        _eligibility(_plant("good"), _plant("bad", structural=True)),
        _policy(),
    )
    states = {item["plant_id"]: item["status"] for item in report["decisions"]}

    assert states == {"good": "ADMITTED", "bad": "REJECT"}
    assert report["counts"] == {"total": 2, "admitted": 1, "hold": 0, "rejected": 1}
    assert report["final_training_selection_ready"] is True
    assert report["service_inventory_filtered"] is False
    assert report["model_population_filtered"] is True
    assert report["next_gate"] == "model_specific_forecast_readiness"


def test_v2_preserves_relational_rejection_reasons_from_eligibility() -> None:
    reasons = [
        "hourly_gap_gt_1h_in_splits:validation",
        "non_train_split_rows_exceed_train:test",
    ]
    report = classify_training_population(
        _eligibility(
            _plant(
                "reject",
                structural=True,
                reasons=reasons,
                rows=(100, 80, 50, 120),
                gaps=(0, 1, 0, 0),
            )
        ),
        _policy(),
    )
    decision = report["decisions"][0]

    assert decision["status"] == "REJECT"
    assert decision["reasons"] == reasons
    assert report["final_training_selection_ready"] is False
    assert report["selection_blockers"] == ["no_admitted_plants"]


def test_v1_eligibility_manifest_must_be_reaudited() -> None:
    with pytest.raises(ValueError, match="rerun audit before freeze"):
        classify_training_population(
            _eligibility(_plant("old"), contract="solar-training-data-eligibility.v1"),
            _policy(),
        )


def test_v2_policy_rejects_invented_fixed_minimum_n() -> None:
    policy = _policy()
    policy["current_4way_extension"]["fixed_minimum_rows_per_split"] = 100

    with pytest.raises(ValueError, match="must not invent a fixed minimum-N rule"):
        classify_training_population(_eligibility(_plant("good")), policy)
