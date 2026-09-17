"""발전량·ASOS 연속 겹침과 분할 충분성을 학습 전에 감사한다."""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from solar_forecast.config_loader import PROJECT_ROOT
from solar_forecast.datasets.repository import DatasetLoadPolicy, DatasetRepository
from solar_forecast.evaluation.experiment_config import load_experiment_config
from solar_forecast.evaluation.temporal_split import TemporalSplitConfig, TemporalSplitter
from solar_forecast.infrastructure.artifact_store import sha256_file, write_json_atomic


ELIGIBILITY_CONTRACT = "solar-training-data-eligibility.v1"
SPLITS = ("train", "validation", "calibration", "test")
WEATHER_COLUMNS = (
    "temperature_c",
    "precipitation_mm",
    "sunshine_hours",
    "solar_irradiance_mj_m2",
    "wind_speed_mps",
    "humidity_pct",
    "total_cloud_cover_tenths",
    "low_mid_cloud_cover_tenths",
)
SEASONS = {
    12: "winter", 1: "winter", 2: "winter",
    3: "spring", 4: "spring", 5: "spring",
    6: "summer", 7: "summer", 8: "summer",
    9: "autumn", 10: "autumn", 11: "autumn",
}


def _truthy(values: pd.Series) -> pd.Series:
    if pd.api.types.is_bool_dtype(values):
        return values.fillna(False)
    return values.astype(str).str.strip().str.lower().isin({"true", "1", "yes"})


def _continuity(times: pd.Series) -> dict[str, object]:
    ordered = pd.DatetimeIndex(times.dropna().drop_duplicates().sort_values())
    if ordered.empty:
        return {
            "start": None, "end": None, "rows": 0, "expected_hourly_rows": 0,
            "hourly_coverage": 0.0, "longest_strict_run_hours": 0,
            "gap_count_gt_1h": 0, "max_gap_hours": None,
        }
    expected = int((ordered[-1] - ordered[0]) / pd.Timedelta(hours=1)) + 1
    if len(ordered) == 1:
        longest = 1
        gaps = pd.Series(dtype="float64")
    else:
        gaps = pd.Series((ordered[1:] - ordered[:-1]) / pd.Timedelta(hours=1), dtype="float64")
        run = 1
        longest = 1
        for gap in gaps:
            if gap == 1:
                run += 1
            else:
                longest = max(longest, run)
                run = 1
        longest = max(longest, run)
    return {
        "start": ordered[0].isoformat(),
        "end": ordered[-1].isoformat(),
        "rows": int(len(ordered)),
        "expected_hourly_rows": expected,
        "hourly_coverage": float(len(ordered) / expected),
        "longest_strict_run_hours": int(longest),
        "gap_count_gt_1h": int((gaps > 1).sum()),
        "max_gap_hours": float(gaps.max()) if len(gaps) else 0.0,
    }


def _partition_summary(frame: pd.DataFrame) -> dict[str, object]:
    if frame.empty:
        return {"rows": 0, "months": [], "seasons": []}
    return {
        "rows": int(len(frame)),
        "months": sorted(frame["timestamp"].dt.strftime("%Y-%m").unique().tolist()),
        "seasons": sorted(frame["timestamp"].dt.month.map(SEASONS).unique().tolist()),
    }


def audit_training_eligibility(
    frame: pd.DataFrame,
    split_config: TemporalSplitConfig,
    *,
    weather_columns: tuple[str, ...] = WEATHER_COLUMNS,
) -> dict[str, object]:
    """Build evidence for plant-period admission without inventing minimum thresholds.

    The result intentionally distinguishes structural rejection from threshold review.
    Final minimum duration/coverage/sample thresholds are chosen only after this report
    is inspected on the real Gold population.
    """
    required = {"timestamp", "plant_id", "generation_mwh", "quality_train_eligible"}
    if missing := required - set(frame.columns):
        raise ValueError(f"Training eligibility requires Gold columns: {sorted(missing)}")
    available_weather = tuple(column for column in weather_columns if column in frame.columns)
    if not available_weather:
        raise ValueError("Training eligibility requires at least one ASOS weather column")

    source = frame.copy()
    source["timestamp"] = TemporalSplitter._timestamps(source["timestamp"])
    source["plant_id"] = source["plant_id"].astype("string")
    identity_ok = source["plant_id"].notna() & source["plant_id"].str.strip().ne("")
    target_ok = np.isfinite(pd.to_numeric(source["generation_mwh"], errors="coerce"))
    quality_ok = _truthy(source["quality_train_eligible"])
    time_ok = source["timestamp"].notna()
    eligible_target = source.loc[identity_ok & target_ok & quality_ok & time_ok].copy()
    if eligible_target.empty:
        raise ValueError("No quality-eligible solar target rows remain for eligibility audit")
    if eligible_target.duplicated(["plant_id", "timestamp"]).any():
        raise ValueError("Duplicate quality-eligible plant-hour keys must be resolved before eligibility audit")

    splitter = TemporalSplitter(split_config)
    boundaries = splitter.boundaries(eligible_target["timestamp"])
    eligible_target["split"] = splitter.labels(eligible_target["timestamp"], boundaries)
    eligible_target["weather_any"] = eligible_target[list(available_weather)].notna().any(axis=1)

    plants: list[dict[str, object]] = []
    for plant_id, plant in eligible_target.groupby("plant_id", observed=True, sort=True):
        weather_overlap = plant.loc[plant["weather_any"]].copy()
        target_continuity = _continuity(plant["timestamp"])
        weather_continuity = _continuity(weather_overlap["timestamp"])
        split_summary: dict[str, object] = {}
        empty_overlap_splits: list[str] = []
        for name in SPLITS:
            target_part = plant.loc[plant["split"].eq(name)]
            overlap_part = weather_overlap.loc[weather_overlap["split"].eq(name)]
            if overlap_part.empty:
                empty_overlap_splits.append(name)
            split_summary[name] = {
                "eligible_target": _partition_summary(target_part),
                "generation_weather_overlap": _partition_summary(overlap_part),
            }

        weather_coverage = {}
        denominator = max(1, len(plant))
        for column in available_weather:
            observed = int(plant[column].notna().sum())
            weather_coverage[column] = {
                "observed_rows": observed,
                "coverage": float(observed / denominator),
            }

        hard_reasons: list[str] = []
        if weather_overlap.empty:
            hard_reasons.append("no_generation_weather_overlap")
        if empty_overlap_splits:
            hard_reasons.append("empty_overlap_splits:" + ",".join(empty_overlap_splits))
        status = "STRUCTURAL_REJECT" if hard_reasons else "CANDIDATE_REQUIRES_THRESHOLD_REVIEW"
        plants.append({
            "plant_id": str(plant_id),
            "status": status,
            "hard_reject_reasons": hard_reasons,
            "target_continuity": target_continuity,
            "generation_weather_overlap": weather_continuity,
            "weather_column_coverage": weather_coverage,
            "splits": split_summary,
        })

    candidates = [item for item in plants if item["status"] == "CANDIDATE_REQUIRES_THRESHOLD_REVIEW"]
    rejected = [item for item in plants if item["status"] == "STRUCTURAL_REJECT"]
    return {
        "contract": ELIGIBILITY_CONTRACT,
        "selection_stage": "evidence_first_before_final_thresholds",
        "fixed_start_year_used": False,
        "weather_columns": list(available_weather),
        "boundaries": boundaries.to_dict(),
        "split_mode": split_config.split_mode,
        "population": {
            "input_rows": int(len(source)),
            "quality_eligible_target_rows": int(len(eligible_target)),
            "plants": int(eligible_target["plant_id"].nunique()),
            "candidate_plants": len(candidates),
            "structurally_rejected_plants": len(rejected),
        },
        "plants": plants,
        "final_thresholds_applied": False,
        "final_training_selection_ready": False,
        "next_decision": (
            "Inspect real overlap/duration/split distributions, then freeze explicit minimum "
            "coverage and effective-sample thresholds before filtering training input."
        ),
        "prediction_or_training_performed": False,
        "test_used_for_selection": False,
    }


def run_training_eligibility_audit(
    config_path: Path,
    *,
    data_path: Path | None = None,
    output_path: Path | None = None,
    project_root: Path = PROJECT_ROOT,
) -> dict[str, object]:
    """Audit current solar Gold and optionally persist an immutable evidence manifest."""
    config_path = config_path if config_path.is_absolute() else project_root / config_path
    values = load_experiment_config(config_path, project_root=project_root)
    source = data_path or Path(values["input_dataset"])
    source = source if source.is_absolute() else project_root / source
    destination = output_path if output_path is None or output_path.is_absolute() else project_root / output_path
    if destination is not None and (
        destination.resolve() == config_path.resolve()
        or destination.resolve() == source.resolve()
        or (source.is_dir() and source.resolve() in destination.resolve().parents)
    ):
        raise ValueError("Eligibility output must not replace configuration or source data")

    files = DatasetRepository._training_files(source)
    inventory = [
        {
            "path": path.relative_to(source).as_posix() if source.is_dir() else path.name,
            "bytes": path.stat().st_size,
            "sha256": sha256_file(path),
        }
        for path in files
    ]
    columns = ["timestamp", "plant_id", "generation_mwh", "quality_train_eligible", *WEATHER_COLUMNS]
    _, frame, loading = DatasetRepository(source.parent).load_training_frame(
        source,
        columns=columns,
        numeric_columns=["generation_mwh", *WEATHER_COLUMNS],
        equals_filters={"energy_source": "solar"},
        policy=DatasetLoadPolicy(numeric_dtype="float64"),
    )
    split_values = dict(values.get("split", {}))
    split_values["purge_gap_hours"] = max(
        int(split_values.get("purge_gap_hours", 168)),
        max(int(value) for value in values["horizons_hours"]),
    )
    report = audit_training_eligibility(frame, TemporalSplitConfig.from_mapping(split_values))
    report.update({
        "source": str(source),
        "source_files": inventory,
        "loading": loading.to_dict(),
        "experiment_config": str(config_path),
        "experiment_config_sha256": sha256_file(config_path),
    })
    if destination is not None:
        write_json_atomic(destination, report)
    return report
