"""학습 없이 고정 시간 분할의 발전소·월·계절 범위와 신규 발전소를 감사."""
from __future__ import annotations

from dataclasses import asdict
from pathlib import Path

import numpy as np
import pandas as pd

from solar_forecast.config_loader import PROJECT_ROOT
from solar_forecast.datasets.repository import DatasetLoadPolicy, DatasetRepository
from solar_forecast.evaluation.experiment_config import load_experiment_config
from solar_forecast.evaluation.temporal_split import TemporalSplitConfig, TemporalSplitter
from solar_forecast.infrastructure.artifact_store import sha256_file, write_json_atomic


SPLIT_AUDIT_CONTRACT = "solar-temporal-split-audit.v1"
SPLIT_NAMES = ("train", "validation", "calibration", "test")
SEASONS = {12: "winter", 1: "winter", 2: "winter", 3: "spring", 4: "spring", 5: "spring",
           6: "summer", 7: "summer", 8: "summer", 9: "autumn", 10: "autumn", 11: "autumn"}


def _coverage(frame: pd.DataFrame) -> dict:
    return {
        "rows": len(frame),
        "plants": int(frame["plant_id"].nunique()),
        "start": frame["timestamp"].min().isoformat() if len(frame) else None,
        "end": frame["timestamp"].max().isoformat() if len(frame) else None,
        "months": sorted(frame["timestamp"].dt.strftime("%Y-%m").unique().tolist()),
        "seasons": sorted(frame["timestamp"].dt.month.map(SEASONS).unique().tolist()),
    }


def audit_temporal_frame(frame: pd.DataFrame, config: TemporalSplitConfig) -> dict:
    """Describe target-row coverage, not model accuracy or complete input windows.

    All historical rows remain available to later model window builders. Purge
    rows and rows beyond test_end are excluded only from partition scoring here.
    """
    required = {"timestamp", "plant_id", "generation_mwh", "quality_train_eligible"}
    if missing := required - set(frame.columns):
        raise ValueError(f"Split audit requires Gold columns: {sorted(missing)}")
    source = frame.copy()
    splitter = TemporalSplitter(config)
    source["timestamp"] = splitter._timestamps(source["timestamp"])
    quality = source["quality_train_eligible"].astype(str).str.strip().str.lower().isin(["true", "1", "yes"])
    target_valid = np.isfinite(pd.to_numeric(source["generation_mwh"], errors="coerce"))
    identity_valid = source["plant_id"].notna() & source["plant_id"].astype(str).str.strip().ne("")
    time_valid = source["timestamp"].notna()
    eligible = source.loc[quality & target_valid & identity_valid & time_valid].copy()
    if eligible.empty:
        raise ValueError("No eligible observed target rows remain for a split audit")
    eligible["plant_id"] = eligible["plant_id"].astype(str)
    if eligible.duplicated(["plant_id", "timestamp"]).any():
        raise ValueError("Duplicate eligible plant-hour keys must be resolved before split auditing")
    boundaries = splitter.boundaries(eligible["timestamp"])
    labels = splitter.labels(eligible["timestamp"], boundaries)
    eligible["split"] = labels
    train_plants = set(eligible.loc[labels.eq("train"), "plant_id"])
    in_partitions = eligible.loc[labels.notna()].copy()
    in_partitions["month"] = in_partitions["timestamp"].dt.strftime("%Y-%m")
    in_partitions["season"] = in_partitions["timestamp"].dt.month.map(SEASONS)
    summaries = {}
    for name in SPLIT_NAMES:
        partition = eligible.loc[labels.eq(name)]
        cold = partition.loc[~partition["plant_id"].isin(train_plants)]
        summaries[name] = {**_coverage(partition), "rows_without_train_history": len(cold),
                           "plants_without_train_history": sorted(cold["plant_id"].unique().tolist())}

    per_plant = []
    for plant_id, plant in eligible.groupby("plant_id", observed=True, sort=True):
        identity = {"plant_id": plant_id, "has_train_history": plant_id in train_plants}
        for field in ("plant", "region"):
            if field in plant:
                identity[field] = str(plant[field].iloc[0])
        for name in SPLIT_NAMES:
            per_plant.append({**identity, "split": name, **_coverage(plant.loc[plant["split"].eq(name)])})

    def grouped_coverage(keys: list[str]) -> list[dict]:
        records = []
        for names, group in in_partitions.groupby(keys, observed=True, sort=True):
            names = names if isinstance(names, tuple) else (names,)
            records.append({**dict(zip(keys, names)), **_coverage(group)})
        return records

    beyond_end = eligible["timestamp"].gt(boundaries.test_end) if boundaries.test_end is not None else pd.Series(False, index=eligible.index)
    warnings = [f"empty_partition:{name}" for name, summary in summaries.items() if not summary["rows"]]
    warnings.extend(f"no_train_history:{name}" for name in SPLIT_NAMES[1:] if summaries[name]["rows_without_train_history"])
    return {
        "contract": SPLIT_AUDIT_CONTRACT,
        "scope": "data_only_eligible_observed_target_rows_before_horizon_lookback_and_feature_gates",
        "prediction_or_training_performed": False,
        "timestamp_convention": "timezone_naive_KST_inclusive_end_timestamp_date_only_means_midnight",
        "split_mode": config.split_mode,
        "split_config": {key: value.isoformat() if isinstance(value, pd.Timestamp) else value for key, value in asdict(config).items()},
        "boundaries": boundaries.to_dict(),
        "population": {
            "input_rows": len(source), "eligible_observed_rows": len(eligible),
            "excluded_rows": len(source) - len(eligible),
            "exclusion_reason_counts_may_overlap": {
                "quality_not_eligible": int((~quality).sum()), "non_finite_target": int((~target_valid).sum()),
                "invalid_timestamp": int((~time_valid).sum()), "missing_plant_id": int((~identity_valid).sum()),
            },
            "assigned_rows": len(in_partitions),
            "purge_gap_rows": int((labels.isna() & ~beyond_end).sum()),
            "after_test_end_rows": int(beyond_end.sum()),
        },
        "splits": summaries,
        "per_plant": per_plant,
        "per_month": grouped_coverage(["split", "month"]),
        "per_season": grouped_coverage(["split", "season"]),
        "per_plant_month": grouped_coverage(["split", "plant_id", "month"]),
        "per_plant_season": grouped_coverage(["split", "plant_id", "season"]),
        "warnings": warnings,
        "limitations": [
            "Counts are observed eligible target rows; final model sample counts may be lower after exact-time history and feature checks.",
            "Has Train history means at least one eligible Train row, not sufficient seasonal coverage or model readiness.",
            "Historical observation timestamps do not prove when providers published those observations.",
            "Rows excluded by the purge gap may still be legitimate past context for a later forecast; this audit never rewrites the dataset.",
        ],
    }


def run_split_audit(
    config_path: Path, *, data_path: Path | None = None, output_path: Path | None = None,
    project_root: Path = PROJECT_ROOT,
) -> dict:
    """Load only Gold metadata/targets and write an auditable data-only report."""
    config_path = config_path if config_path.is_absolute() else project_root / config_path
    values = load_experiment_config(config_path, project_root=project_root)
    source = data_path or Path(values["input_dataset"])
    source = source if source.is_absolute() else project_root / source
    destination = output_path if output_path is None or output_path.is_absolute() else project_root / output_path
    if destination is not None and destination.resolve() == config_path.resolve():
        raise ValueError("Split audit output must differ from the experiment configuration")
    if destination is not None and (destination.resolve() == source.resolve() or
                                    (source.is_dir() and source.resolve() in destination.resolve().parents)):
        raise ValueError("Split audit output must be outside the source dataset")
    files = DatasetRepository._training_files(source)
    inventory = [{"path": path.relative_to(source).as_posix() if source.is_dir() else path.name,
                  "bytes": path.stat().st_size, "sha256": sha256_file(path)} for path in files]
    _, frame, loading = DatasetRepository(source.parent).load_training_frame(
        source, columns=["timestamp", "plant_id", "plant", "region", "generation_mwh", "quality_train_eligible"],
        numeric_columns=["generation_mwh"], equals_filters={"energy_source": "solar"},
        policy=DatasetLoadPolicy(numeric_dtype="float64"),
    )
    split_values = dict(values.get("split", {}))
    # Candidate construction protects every model with at least its forecast
    # horizon. Audit each effective gap when horizons exceed the configured gap.
    configured_gap = split_values.get("purge_gap_hours", 168)
    horizons_by_gap: dict[int, list[int]] = {}
    for horizon in values["horizons_hours"]:
        horizons_by_gap.setdefault(max(configured_gap, horizon), []).append(horizon)
    audits = []
    for gap, horizons in sorted(horizons_by_gap.items()):
        config = TemporalSplitConfig.from_mapping({**split_values, "purge_gap_hours": gap})
        audits.append({"horizons_hours": horizons, **audit_temporal_frame(frame, config)})
    report = {
        "contract": SPLIT_AUDIT_CONTRACT, "source": str(source),
        "experiment_config": str(config_path), "experiment_config_sha256": sha256_file(config_path), "source_files": inventory,
        "loading": loading.to_dict(), "audits": audits,
        "prediction_or_training_performed": False,
    }
    if destination is not None:
        write_json_atomic(destination, report)
    return report
