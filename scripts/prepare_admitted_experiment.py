"""Create a run-local experiment config using the frozen ADMITTED Solar dataset."""
from __future__ import annotations

import argparse
import json
from pathlib import Path


EXPECTED_POPULATION_CONTRACT = "solar-training-population.v2"
EXPECTED_POLICY_CONTRACT = "solar-training-admission-policy.v2"
EXPECTED_DATASET_CONTRACT = "solar-admitted-training-dataset.v1"


def _read(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", default="config/experiments/optimized.json")
    parser.add_argument("--population", default="artifacts/evaluation/training_population.json")
    parser.add_argument("--policy", default="config/training_admission.json")
    parser.add_argument("--data", default="artifacts/datasets/admitted_training.csv.gz")
    parser.add_argument("--data-manifest", default="artifacts/datasets/admitted_training.csv.gz.manifest.json")
    parser.add_argument("--output", default="artifacts/evaluation/admitted_experiment.json")
    args = parser.parse_args()

    config_path = Path(args.config)
    population_path = Path(args.population)
    policy_path = Path(args.policy)
    data_path = Path(args.data)
    data_manifest_path = Path(args.data_manifest)
    output_path = Path(args.output)

    config = _read(config_path)
    population = _read(population_path)
    policy = _read(policy_path)
    data_manifest = _read(data_manifest_path)

    if population.get("contract") != EXPECTED_POPULATION_CONTRACT:
        raise SystemExit(
            f"Expected {EXPECTED_POPULATION_CONTRACT}; rerun eligibility audit and population freeze"
        )
    if policy.get("contract") != EXPECTED_POLICY_CONTRACT or policy.get("status") != "frozen":
        raise SystemExit("Training admission policy v2 must be frozen")
    if data_manifest.get("contract") != EXPECTED_DATASET_CONTRACT:
        raise SystemExit(
            f"Expected {EXPECTED_DATASET_CONTRACT}; materialize the admitted dataset first"
        )
    if not data_path.is_file():
        raise SystemExit(f"Materialized admitted dataset not found: {data_path}")
    if not population.get("final_training_selection_ready"):
        raise SystemExit(
            "Training population is not ready: "
            + ",".join(population.get("selection_blockers") or ["unknown_blocker"])
        )
    if population.get("service_inventory_filtered") is not False:
        raise SystemExit("Population contract must not filter the service inventory")
    if data_manifest.get("service_inventory_filtered") is not False:
        raise SystemExit("Materialized dataset must not redefine service inventory population")

    admitted = [str(value) for value in population.get("admitted_plant_ids") or []]
    materialized = [str(value) for value in data_manifest.get("admitted_plant_ids") or []]
    if not admitted:
        raise SystemExit("No ADMITTED plants are available for model training")
    if sorted(admitted) != sorted(materialized):
        raise SystemExit("Materialized dataset plant IDs differ from frozen training population")

    resolved = dict(config)
    resolved["input_dataset"] = str(data_path)
    resolved["admitted_plant_ids"] = admitted
    resolved["training_population_manifest"] = str(population_path)
    resolved["training_population_contract"] = EXPECTED_POPULATION_CONTRACT
    resolved["training_population_policy_version"] = population.get("policy_version")
    resolved["admitted_dataset_manifest"] = str(data_manifest_path)
    resolved["admitted_dataset_contract"] = EXPECTED_DATASET_CONTRACT
    resolved["service_inventory_population_filtered"] = False

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = output_path.with_suffix(output_path.suffix + ".part")
    temporary.write_text(
        json.dumps(resolved, ensure_ascii=False, indent=2, allow_nan=False) + "\n",
        encoding="utf-8",
    )
    temporary.replace(output_path)

    print(f"admitted_plants={len(admitted)}")
    print(f"input_dataset={data_path}")
    print("service_inventory_filtered=False")
    print(f"resolved_config={output_path}")
    print("next=Run local preflight/readiness with this resolved config, then start GPU training only after it passes.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
