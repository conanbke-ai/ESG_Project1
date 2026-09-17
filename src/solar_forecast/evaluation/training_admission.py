"""Freeze the Solar AI training population from eligibility.v2 evidence."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any


POLICY_CONTRACT = "solar-training-admission-policy.v2"
POPULATION_CONTRACT = "solar-training-population.v2"
ELIGIBILITY_CONTRACT = "solar-training-data-eligibility.v2"
SPLITS = ("train", "validation", "calibration", "test")


@dataclass(frozen=True)
class AdmissionDecision:
    plant_id: str
    status: str
    reasons: tuple[str, ...]
    metrics: dict[str, float | int]

    def to_dict(self) -> dict[str, Any]:
        return {
            "plant_id": self.plant_id,
            "status": self.status,
            "reasons": list(self.reasons),
            "metrics": self.metrics,
        }


def validate_admission_policy(policy: dict[str, Any]) -> None:
    if policy.get("contract") != POLICY_CONTRACT:
        raise ValueError(f"Admission policy requires contract {POLICY_CONTRACT}")
    if policy.get("status") != "frozen":
        raise ValueError("V2 recovered relational admission policy must be frozen")
    if policy.get("fixed_start_year") not in {None, False}:
        raise ValueError("V2 admission policy does not allow a fixed start year")
    extension = policy.get("current_4way_extension")
    if not isinstance(extension, dict):
        raise ValueError("V2 admission policy requires current_4way_extension")
    required_true = (
        "require_all_splits_present",
        "require_validation_rows_lte_train_rows",
        "require_calibration_rows_lte_train_rows",
        "require_test_rows_lte_train_rows",
    )
    if any(extension.get(key) is not True for key in required_true):
        raise ValueError("V2 admission policy must preserve recovered relational rules")
    if int(extension.get("reject_split_timestamp_gap_gt_hours", -1)) != 1:
        raise ValueError("V2 admission policy requires rejection of >1h split gaps")
    if extension.get("fixed_minimum_rows_per_split") is not None:
        raise ValueError("V2 policy must not invent a fixed minimum-N rule")


def _metrics(item: dict[str, Any]) -> dict[str, float | int]:
    overlap = item["generation_weather_overlap"]
    weather_values = [
        float(value["coverage"])
        for value in item.get("weather_column_coverage", {}).values()
    ]
    values: dict[str, float | int] = {
        "overlap_hourly_coverage": float(overlap["hourly_coverage"]),
        "longest_strict_run_hours": int(overlap["longest_strict_run_hours"]),
        "weather_column_coverage_min": min(weather_values) if weather_values else 0.0,
    }
    for split in SPLITS:
        summary = item["splits"][split]["generation_weather_overlap"]
        values[f"{split}_rows"] = int(summary["rows"])
        continuity = summary.get("continuity", {})
        values[f"{split}_gap_count_gt_1h"] = int(
            continuity.get("gap_count_gt_1h", 0)
        )
    return values


def classify_training_population(
    eligibility: dict[str, Any], policy: dict[str, Any]
) -> dict[str, Any]:
    """Classify only the model population; never the nationwide service inventory."""
    if eligibility.get("contract") != ELIGIBILITY_CONTRACT:
        raise ValueError(
            f"Eligibility manifest requires contract {ELIGIBILITY_CONTRACT}; rerun audit before freeze"
        )
    validate_admission_policy(policy)

    recovered = eligibility.get("historical_rule_recovery", {})
    if recovered.get("fixed_minimum_rows_per_split_found") is not False:
        raise ValueError("Eligibility v2 must state that no fixed minimum-N evidence was found")
    if recovered.get("historical_3way_extended_to_current_4way") is not True:
        raise ValueError("Eligibility v2 must apply the recovered rule to the current four-way split")

    decisions: list[AdmissionDecision] = []
    for item in eligibility.get("plants", []):
        plant_id = str(item["plant_id"])
        metrics = _metrics(item)
        structural_reasons = tuple(item.get("hard_reject_reasons", []))
        if item.get("status") == "STRUCTURAL_REJECT" or structural_reasons:
            decisions.append(
                AdmissionDecision(
                    plant_id=plant_id,
                    status="REJECT",
                    reasons=structural_reasons or ("structural_reject",),
                    metrics=metrics,
                )
            )
            continue

        decisions.append(
            AdmissionDecision(
                plant_id=plant_id,
                status="ADMITTED",
                reasons=("recovered_relational_rules_passed",),
                metrics=metrics,
            )
        )

    admitted = [item.plant_id for item in decisions if item.status == "ADMITTED"]
    rejected = [item.plant_id for item in decisions if item.status == "REJECT"]
    return {
        "contract": POPULATION_CONTRACT,
        "policy_version": policy.get("version"),
        "policy_status": policy["status"],
        "fixed_start_year_used": False,
        "service_inventory_filtered": False,
        "model_population_filtered": True,
        "decisions": [item.to_dict() for item in decisions],
        "counts": {
            "total": len(decisions),
            "admitted": len(admitted),
            "hold": 0,
            "rejected": len(rejected),
        },
        "admitted_plant_ids": admitted,
        "hold_plant_ids": [],
        "rejected_plant_ids": rejected,
        "final_training_selection_ready": bool(admitted),
        "selection_blockers": [] if admitted else ["no_admitted_plants"],
        "next_gate": "model_specific_forecast_readiness",
    }
