"""Freeze the Solar training population from an evidence manifest and explicit policy."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any


POLICY_CONTRACT = "solar-training-admission-policy.v1"
POPULATION_CONTRACT = "solar-training-population.v1"
ELIGIBILITY_CONTRACT = "solar-training-data-eligibility.v1"
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


def _require_thresholds(policy: dict[str, Any], name: str) -> dict[str, Any]:
    block = policy.get(name)
    if not isinstance(block, dict):
        raise ValueError(f"{name} must be an object")
    if policy.get("status") == "frozen":
        required = (
            "min_overlap_hourly_coverage",
            "min_longest_strict_run_hours",
            "min_weather_column_coverage",
        )
        if any(block.get(key) is None for key in required):
            raise ValueError(f"Frozen policy requires every {name} scalar threshold")
        split = block.get("min_split_rows")
        if not isinstance(split, dict) or any(split.get(key) is None for key in SPLITS):
            raise ValueError(f"Frozen policy requires every {name}.min_split_rows threshold")
    return block


def validate_admission_policy(policy: dict[str, Any]) -> None:
    if policy.get("contract") != POLICY_CONTRACT:
        raise ValueError(f"Admission policy requires contract {POLICY_CONTRACT}")
    if policy.get("status") not in {"draft", "frozen"}:
        raise ValueError("Admission policy status must be draft or frozen")
    if policy.get("fixed_start_year") not in {None, False}:
        raise ValueError("V1 admission policy does not allow a fixed start year")
    admit = _require_thresholds(policy, "admit_thresholds")
    reject = _require_thresholds(policy, "reject_thresholds")
    if policy.get("status") != "frozen":
        return

    scalar_keys = (
        "min_overlap_hourly_coverage",
        "min_longest_strict_run_hours",
        "min_weather_column_coverage",
    )
    for key in scalar_keys:
        a = float(admit[key])
        r = float(reject[key])
        if a < 0 or r < 0 or r > a:
            raise ValueError(f"Thresholds must satisfy 0 <= reject <= admit for {key}")
        if "coverage" in key and (a > 1 or r > 1):
            raise ValueError(f"Coverage threshold must be <= 1 for {key}")
    for split in SPLITS:
        a = int(admit["min_split_rows"][split])
        r = int(reject["min_split_rows"][split])
        if a < 0 or r < 0 or r > a:
            raise ValueError(
                f"Split thresholds must satisfy 0 <= reject <= admit for {split}"
            )


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
        values[f"{split}_rows"] = int(
            item["splits"][split]["generation_weather_overlap"]["rows"]
        )
    return values


def _threshold_failures(
    metrics: dict[str, float | int], thresholds: dict[str, Any], prefix: str
) -> list[str]:
    checks = (
        (
            "overlap_hourly_coverage",
            float(thresholds["min_overlap_hourly_coverage"]),
        ),
        (
            "longest_strict_run_hours",
            int(thresholds["min_longest_strict_run_hours"]),
        ),
        (
            "weather_column_coverage_min",
            float(thresholds["min_weather_column_coverage"]),
        ),
    )
    reasons = [
        f"{prefix}:{name}<{minimum}"
        for name, minimum in checks
        if float(metrics[name]) < float(minimum)
    ]
    for split in SPLITS:
        minimum = int(thresholds["min_split_rows"][split])
        if int(metrics[f"{split}_rows"]) < minimum:
            reasons.append(f"{prefix}:{split}_rows<{minimum}")
    return reasons


def classify_training_population(
    eligibility: dict[str, Any], policy: dict[str, Any]
) -> dict[str, Any]:
    if eligibility.get("contract") != ELIGIBILITY_CONTRACT:
        raise ValueError(f"Eligibility manifest requires contract {ELIGIBILITY_CONTRACT}")
    validate_admission_policy(policy)

    decisions: list[AdmissionDecision] = []
    for item in eligibility.get("plants", []):
        plant_id = str(item["plant_id"])
        metrics = _metrics(item)
        structural_reasons = tuple(item.get("hard_reject_reasons", []))
        if item.get("status") == "STRUCTURAL_REJECT" or structural_reasons:
            decisions.append(
                AdmissionDecision(
                    plant_id,
                    "REJECT",
                    structural_reasons or ("structural_reject",),
                    metrics,
                )
            )
            continue
        if policy["status"] != "frozen":
            decisions.append(
                AdmissionDecision(plant_id, "HOLD", ("admission_policy_not_frozen",), metrics)
            )
            continue

        reject_reasons = _threshold_failures(
            metrics, policy["reject_thresholds"], "reject_floor"
        )
        if reject_reasons:
            decisions.append(
                AdmissionDecision(plant_id, "REJECT", tuple(reject_reasons), metrics)
            )
            continue
        hold_reasons = _threshold_failures(
            metrics, policy["admit_thresholds"], "admit_gate"
        )
        if hold_reasons:
            decisions.append(
                AdmissionDecision(plant_id, "HOLD", tuple(hold_reasons), metrics)
            )
            continue
        decisions.append(AdmissionDecision(plant_id, "ADMITTED", ("all_gates_passed",), metrics))

    admitted = [item.plant_id for item in decisions if item.status == "ADMITTED"]
    hold = [item.plant_id for item in decisions if item.status == "HOLD"]
    rejected = [item.plant_id for item in decisions if item.status == "REJECT"]
    ready = policy["status"] == "frozen" and bool(admitted) and not hold
    return {
        "contract": POPULATION_CONTRACT,
        "policy_version": policy.get("version"),
        "policy_status": policy["status"],
        "fixed_start_year_used": False,
        "decisions": [item.to_dict() for item in decisions],
        "counts": {
            "total": len(decisions),
            "admitted": len(admitted),
            "hold": len(hold),
            "rejected": len(rejected),
        },
        "admitted_plant_ids": admitted,
        "hold_plant_ids": hold,
        "rejected_plant_ids": rejected,
        "final_training_selection_ready": ready,
        "selection_blockers": (
            []
            if ready
            else [
                reason
                for reason, active in (
                    ("admission_policy_not_frozen", policy["status"] != "frozen"),
                    ("no_admitted_plants", not admitted),
                    ("hold_plants_remain", bool(hold)),
                )
                if active
            ]
        ),
    }
