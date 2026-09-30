"""Materialize the frozen ADMITTED Solar model population without touching service inventory data."""
from __future__ import annotations

import argparse
from hashlib import sha256
import json
from pathlib import Path

import pandas as pd


POPULATION_CONTRACT = "solar-training-population.v2"
DATASET_CONTRACT = "solar-admitted-training-dataset.v1"


def _truthy(values: pd.Series) -> pd.Series:
    if pd.api.types.is_bool_dtype(values):
        return values.fillna(False)
    return values.astype(str).str.strip().str.lower().isin({"true", "1", "yes"})


def _training_files(source: Path) -> list[Path]:
    if source.is_file():
        return [source]
    if source.is_dir():
        files = sorted(
            path for path in source.rglob("*")
            if path.is_file() and path.name.lower().endswith((".csv", ".csv.gz"))
        )
        if files:
            return files
    raise FileNotFoundError(f"Training source not found: {source}")


def _file_sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", default="file/standardized/model_ready_parts")
    parser.add_argument("--population", default="artifacts/evaluation/training_population.json")
    parser.add_argument("--output", default="artifacts/datasets/admitted_training.csv.gz")
    parser.add_argument("--chunk-rows", type=int, default=100_000)
    args = parser.parse_args()

    source = Path(args.source)
    population_path = Path(args.population)
    output = Path(args.output)
    manifest_path = output.with_suffix(output.suffix + ".manifest.json")

    population = json.loads(population_path.read_text(encoding="utf-8"))
    if population.get("contract") != POPULATION_CONTRACT:
        raise SystemExit(f"Expected {POPULATION_CONTRACT}; rerun audit/freeze first")
    if not population.get("final_training_selection_ready"):
        raise SystemExit("Training population is not ready")
    if population.get("service_inventory_filtered") is not False:
        raise SystemExit("Model selection must not filter service inventory population")

    admitted = {str(value) for value in population.get("admitted_plant_ids") or []}
    if not admitted:
        raise SystemExit("No ADMITTED plant IDs")

    files = _training_files(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(output.name + ".part")
    temporary.unlink(missing_ok=True)

    input_rows = 0
    retained_rows = 0
    written_header = False
    retained_plants: set[str] = set()
    columns_seen: list[str] | None = None
    source_inventory: list[dict[str, object]] = []

    for path in files:
        source_inventory.append({
            "path": path.as_posix(),
            "bytes": path.stat().st_size,
            "sha256": _file_sha256(path),
        })
        for chunk in pd.read_csv(path, chunksize=args.chunk_rows, low_memory=False):
            input_rows += len(chunk)
            required = {"plant_id", "energy_source", "quality_train_eligible"}
            missing = required - set(chunk.columns)
            if missing:
                raise ValueError(f"Source partition missing required columns {sorted(missing)}: {path}")
            if columns_seen is None:
                columns_seen = list(chunk.columns)
            elif list(chunk.columns) != columns_seen:
                raise ValueError(f"Source partition schema mismatch: {path}")

            mask = (
                chunk["plant_id"].astype(str).isin(admitted)
                & chunk["energy_source"].astype(str).eq("solar")
                & _truthy(chunk["quality_train_eligible"])
            )
            selected = chunk.loc[mask].copy()
            if selected.empty:
                continue
            selected.to_csv(
                temporary,
                mode="w" if not written_header else "a",
                header=not written_header,
                index=False,
                encoding="utf-8-sig" if not written_header else "utf-8",
                compression="gzip",
            )
            written_header = True
            retained_rows += len(selected)
            retained_plants.update(selected["plant_id"].astype(str).unique())

    if not written_header or retained_rows == 0:
        temporary.unlink(missing_ok=True)
        raise SystemExit("No rows remained after applying the frozen model population")
    if retained_plants != admitted:
        temporary.unlink(missing_ok=True)
        missing = sorted(admitted - retained_plants)
        extra = sorted(retained_plants - admitted)
        raise SystemExit(f"Materialized plant set mismatch missing={missing} extra={extra}")

    temporary.replace(output)
    payload = {
        "contract": DATASET_CONTRACT,
        "source": str(source),
        "source_files": source_inventory,
        "training_population_manifest": str(population_path),
        "training_population_contract": POPULATION_CONTRACT,
        "admitted_plant_ids": sorted(admitted),
        "service_inventory_filtered": False,
        "filters": {
            "plant_id": "frozen ADMITTED allow-list",
            "energy_source": "solar",
            "quality_train_eligible": True,
        },
        "input_rows_scanned": input_rows,
        "retained_rows": retained_rows,
        "retained_plants": len(retained_plants),
        "output": str(output),
        "output_sha256": _file_sha256(output),
    }
    manifest_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False) + "\n",
        encoding="utf-8",
    )

    print(f"input_rows_scanned={input_rows}")
    print(f"retained_rows={retained_rows}")
    print(f"retained_plants={len(retained_plants)}")
    print("service_inventory_filtered=False")
    print(f"dataset={output}")
    print(f"manifest={manifest_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
