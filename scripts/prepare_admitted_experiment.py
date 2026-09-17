"""Create a run-local experiment config using only frozen ADMITTED plant IDs."""
from __future__ import annotations

import argparse
import json
from pathlib import Path


EXPECTED_POPULATION_CONTRACT = "solar-training-population.v2"
EXPECTED_POLICY_CONTRACT = "solar-training-admission-policy.v2"


def _read(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", default="config/experiments/optimized.json")
    parser.add_argument("--population", default="artifacts/evaluation/training_population.json")
    parser.add_argument("--policy", default="config/training_admission.json")
    parser.add_argument("--output", default="artifacts/evaluation/admitted_experiment.json")
    args = parser.parse_args()

    config_path = Path(args.config)
    population_path = Path(args.population)
    policy_path = Path(args.policy)
    output_path = Path(args.output)

    config = _read(config_path)
    population = _read(population_path)
    policy = _read(policy_path)

    if population.get("contract") != EXPECTED_POPULATION_CONTRACT:
        raise SystemExit(
            f"Expected {EXPECTED_POPULATION_CONTRACT}; rerun eligibility audit and population freeze"
        )
    if policy.get("contract") != EXPECTED_POLICY_CONTRACT or policy.get("status") != "frozen":
        raise SystemExit("Training admission policy v2 must be frozen")
    if not population.get("final_training_selection_ready"):
        raise SystemExit(
            "Training population is not ready: "
            + ",".join(population.get("selection_blockers") or ["unknown_blocker"])
        )
    if population.get("service_inventory_filtered") is not False:
        raise SystemExit("Population contract must not filter the service inventory")

    admitted = [str(value) for value in population.get("admitted_plant_ids") or []]
    if not admitted:
        raise SystemExit("No ADMITTED plants are available for model training")

    resolved = dict(config)
    resolved["admitted_plant_ids"] = admitted
    resolved["training_population_manifest"] = str(population_path)
    resolved["training_population_contract"] = EXPECTED_POPULATION_CONTRACT
    resolved["training_population_policy_version"] = population.get("policy_version")
    resolved["service_inventory_population_filtered"] = False

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = output_path.with_suffix(output_path.suffix + ".part")
    temporary.write_text(
        json.dumps(resolved, ensure_ascii=False, indent=2, allow_nan=False) + "\n",
        encoding="utf-8",
    )
    temporary.replace(output_path)

    print(f"admitted_plants={len(admitted)}")
    print("service_inventory_filtered=False")
    print(f"resolved_config={output_path}")
    print("next=Run forecast readiness/preflight with this resolved config before local GPU training.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
