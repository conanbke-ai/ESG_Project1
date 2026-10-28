"""학습 없이 실제 예측 기준시각·연속 입력창·Train 피처의 사용 가능 표본 감사."""
from __future__ import annotations

from dataclasses import dataclass
import hashlib
import json
from pathlib import Path

import numpy as np
import pandas as pd

from solar_forecast.config_loader import ModelJobConfig, PROJECT_ROOT
from solar_forecast.datasets.numeric_preprocessor import NumericPreprocessor, require_model_quality_filter
from solar_forecast.datasets.repository import DatasetLoadPolicy, DatasetRepository
from solar_forecast.evaluation.experiment_config import build_candidate_configs, load_experiment_config
from solar_forecast.evaluation.forecast_samples import build_forecast_samples, forecast_window_positions, validate_observation_frame
from solar_forecast.evaluation.temporal_split import TemporalSplitConfig, TemporalSplitter
from solar_forecast.infrastructure.artifact_store import sha256_file, write_json_atomic


READINESS_CONTRACT = "solar-forecast-readiness.v1"
SPLITS = ("train", "validation", "calibration", "test")


@dataclass
class _SampleGeometry:
    summary: dict
    keys: dict[str, pd.MultiIndex]
    train_input_rows: np.ndarray


def _geometry(frame: pd.DataFrame, config: ModelJobConfig) -> _SampleGeometry:
    values = config.values
    horizon = values["forecast_horizon_hours"]
    length = int(values.get("sequence_length", 168)) if config.model == "cnn_bilstm" else 1
    splitter = TemporalSplitter(TemporalSplitConfig.from_mapping({
        **values, "purge_gap_hours": max(int(values.get("purge_gap_hours", 168)), horizon),
    }))
    boundaries = splitter.boundaries(frame["timestamp"])
    labels = splitter.labels(frame["timestamp"], boundaries)
    target_dtype = np.float32 if config.model == "cnn_bilstm" else None
    finite_target = np.isfinite(frame["generation_mwh"].to_numpy(dtype=target_dtype))
    origin_positions = np.full(len(frame), -1, dtype=np.int64)
    window_valid = np.zeros(len(frame), dtype=bool)
    enough_history = np.zeros(len(frame), dtype=bool)
    for _, group in frame.groupby("plant_id", sort=False, observed=True):
        rows = group.index.to_numpy()
        times = pd.DatetimeIndex(group["timestamp"])
        origins = times.get_indexer(times - pd.Timedelta(hours=horizon))
        found = origins >= 0
        origin_positions[rows[found]] = rows[origins[found]]
        enough_history[rows] = origins >= length - 1
        if config.model == "cnn_bilstm":
            targets, _ = forecast_window_positions(times, horizon_hours=horizon, sequence_length=length)
            window_valid[rows[targets]] = True
    found = origin_positions >= 0
    finite_origin = np.zeros(len(frame), dtype=bool)
    finite_origin[found] = finite_target[origin_positions[found]]
    paired = finite_target & found & finite_origin
    if config.model == "xgboost":
        # Reuse the trainer's exact-time join. Feature values never gate rows;
        # its origin feature positions below produce the same Train inputs.
        metadata = frame[["plant_id", "timestamp", "generation_mwh"]].assign(audit_row=np.arange(len(frame)))
        paired_rows = build_forecast_samples(metadata, [], "generation_mwh", horizon)["audit_row"].to_numpy()
        window_valid[paired_rows] = True
    valid = paired & window_valid
    train_targets = np.flatnonzero(valid & labels.eq("train").fillna(False).to_numpy(dtype=bool))
    train_inputs = np.zeros(len(frame), dtype=bool)
    if config.model == "xgboost":
        train_inputs[origin_positions[train_targets]] = True
    else:
        # Observations are contiguous by plant. Difference-array union counts
        # each input row once, exactly as CNN's Train-fitted preprocessing does.
        usage = np.zeros(len(frame) + 1, dtype=np.int64)
        origins = origin_positions[train_targets]
        np.add.at(usage, origins - length + 1, 1)
        np.add.at(usage, origins + 1, -1)
        train_inputs = np.cumsum(usage[:-1]) > 0
    losses = {
        "non_finite_target": ~finite_target,
        "unavailable_exact_origin": finite_target & ~found,
        "non_finite_origin": finite_target & found & ~finite_origin,
        "insufficient_history": paired & ~enough_history,
        "discontinuous_history": paired & enough_history & ~window_valid,
    }
    summaries, keys = {}, {}
    for name in SPLITS:
        assigned = labels.eq(name).fillna(False).to_numpy(dtype=bool)
        selected = frame.loc[assigned & valid, ["plant_id", "timestamp"]]
        keys[name] = pd.MultiIndex.from_frame(selected)
        summaries[name] = {
            "eligible_target_rows": int(assigned.sum()), "usable_samples": len(selected),
            "losses": {reason: int((assigned & mask).sum()) for reason, mask in losses.items()},
            "per_plant": [
                {"plant_id": str(plant), "eligible_target_rows": int(count),
                 "usable_samples": int(selected["plant_id"].eq(plant).sum())}
                for plant, count in frame.loc[assigned].groupby("plant_id", observed=True).size().reindex(
                    frame["plant_id"].drop_duplicates(), fill_value=0).items()
            ],
        }
    empty = [name for name in SPLITS if not summaries[name]["usable_samples"]]
    if empty:
        raise ValueError(f"{config.model} horizon={horizon} lookback={length}: empty sample partitions {empty}")
    return _SampleGeometry({"boundaries": boundaries.to_dict(), "splits": summaries,
                            "unassigned_target_rows": int(labels.isna().sum())}, keys, train_inputs)


def _feature_sanity(frame: pd.DataFrame, rows: np.ndarray, features: list[str], *, cnn: bool) -> dict:
    inputs = frame.loc[rows, features]
    # CNN materializes its feature series as float32 even with a float64 loader.
    if cnn:
        inputs = inputs.astype("float32")
    inputs = inputs.replace([np.inf, -np.inf], np.nan)
    observations = inputs.notna().sum()
    missing = inputs.isna().sum()
    unique = inputs.nunique(dropna=True)
    return {
        "fit_scope": "unique_observation_rows_actually_used_by_train_samples",
        "rows": len(inputs), "feature_order": features,
        "features": [{"name": name, "observed_rows": int(observations[name]),
                      "missing_rows": int(missing[name]), "missing_fraction": float(missing[name] / len(inputs)),
                      "all_missing": bool(observations[name] == 0),
                      "constant_observed": bool(unique[name] == 1)} for name in features],
        "all_missing_features": [name for name in features if observations[name] == 0],
        "constant_observed_features": [name for name in features if unique[name] == 1],
        "statistics_or_imputation_fitted": False,
        "policy": "cnn_train_median_zero_fallback_and_saved_scaling" if cnn else "xgboost_native_missing_values",
    }


def audit_forecast_frame(frame: pd.DataFrame, candidates: list[tuple[str, ModelJobConfig]], *, minimum_coverage: float) -> dict:
    """Audit one horizon on the trainer-filtered, dtype-converted observation frame.

    Callers must apply DatasetRepository's solar/quality filters first. The file
    runner below is the public loading path. Feature sets share sample geometry;
    missing input features are retained according to the existing model policy.
    """
    if not candidates or not 0 < minimum_coverage <= 1:
        raise ValueError("Candidates and a common coverage threshold in (0, 1] are required")
    horizons = {cfg.values["forecast_horizon_hours"] for _, cfg in candidates}
    if len(horizons) != 1:
        raise ValueError("One readiness cohort must contain one forecast horizon")
    for _, cfg in candidates:
        if cfg.model not in {"xgboost", "cnn_bilstm"} or cfg.values.get("prediction_task") != "historical_forecast":
            raise ValueError("Readiness supports historical XGBoost/CNN-BiLSTM candidates only")
        if "generation_mwh" in cfg.values["feature_columns"]:
            raise ValueError("The future target cannot also be an unshifted feature column")
        if cfg.values.get("target_column") != "generation_mwh":
            raise ValueError("Readiness requires the observed generation_mwh target")
        if cfg.model == "cnn_bilstm" and (cfg.values.get("entity_column") != "plant_id" or cfg.values.get("timestamp_column") != "timestamp"):
            raise ValueError("Readiness requires plant_id/timestamp CNN identity")
    features = list(dict.fromkeys(column for _, cfg in candidates for column in cfg.values["feature_columns"]))
    prepared = NumericPreprocessor(fill_missing=False).transform(
        frame, "generation_mwh", features, passthrough_columns=["plant_id", "timestamp"])
    observations = validate_observation_frame(prepared.frame, entity_column="plant_id", timestamp_column="timestamp").reset_index(drop=True)
    # The common temporal splitter rejects timezone-aware values explicitly.
    TemporalSplitter._timestamps(observations["timestamp"])
    if observations.empty:
        raise ValueError("No observations remain after numeric preprocessing")
    cache: dict[str, _SampleGeometry] = {}
    geometries, results = {}, []
    for candidate_id, cfg in candidates:
        label = f"{cfg.model}:{candidate_id}"
        if label in geometries:
            raise ValueError(f"Duplicate readiness candidate: {label}")
        split = TemporalSplitConfig.from_mapping(cfg.values)
        key = repr((cfg.model, cfg.values.get("sequence_length") if cfg.model == "cnn_bilstm" else 1, split))
        if key not in cache:
            cache[key] = _geometry(observations, cfg)
        geometry = cache[key]
        geometries[label] = geometry
        sanity = _feature_sanity(observations, geometry.train_input_rows, cfg.values["feature_columns"], cnn=cfg.model == "cnn_bilstm")
        if (cfg.model == "cnn_bilstm" and sanity["all_missing_features"]
                and not bool(cfg.values.get("append_missing_indicators", True))):
            raise ValueError(
                "CNN Train features with no observed values require missing indicators: "
                f"{sanity['all_missing_features']}"
            )
        results.append({"model": cfg.model, "candidate_id": candidate_id,
                        "resolved_config_sha256": hashlib.sha256(json.dumps(cfg.values, sort_keys=True).encode()).hexdigest(),
                        "horizon_hours": next(iter(horizons)),
                        "sequence_length": cfg.values.get("sequence_length") if cfg.model == "cnn_bilstm" else None,
                        **geometry.summary, "train_feature_sanity": sanity,
                        "warnings": [f"all_missing_train_feature:{name}" for name in sanity["all_missing_features"]]})
    coverage = {}
    for split in SPLITS:
        indices = [geometry.keys[split] for geometry in cache.values()]
        common, union = indices[0], indices[0]
        for index in indices[1:]:
            common, union = common.intersection(index), union.union(index)
        fraction = len(common) / len(union)
        coverage[split] = {"common_rows": len(common), "union_rows": len(union), "common_fraction": fraction,
                           "minimum_required": minimum_coverage, "meets_threshold": fraction >= minimum_coverage,
                           "candidate_rows": {label: len(value.keys[split]) for label, value in geometries.items()}}
    return {"horizon_hours": next(iter(horizons)), "input_rows": len(frame),
            "missing_target_rows_removed": prepared.dropped_rows,
            "candidates": results, "common_coverage": coverage,
            "coverage_gate_passed": all(coverage[name]["meets_threshold"] for name in SPLITS[1:]),
            "coverage_scope": "all_candidates; validation is required by benchmark; calibration/test is conservative before base selection",
            "prediction_key_equivalence": "plant_id,target_timestamp imply origin=target-horizon for this fixed horizon cohort"}


def run_forecast_readiness(config_path: Path, *, data_path: Path | None = None,
                           output_path: Path | None = None, project_root: Path = PROJECT_ROOT) -> dict:
    """Read Gold with actual trainer filters and report data readiness without ML frameworks."""
    config_path = config_path if config_path.is_absolute() else project_root / config_path
    config_hash = sha256_file(config_path)
    values = load_experiment_config(config_path, project_root=project_root)
    source = data_path or Path(values["input_dataset"])
    source = source if source.is_absolute() else project_root / source
    destination = output_path if output_path is None or output_path.is_absolute() else project_root / output_path
    if destination is not None and (destination.resolve() == config_path.resolve() or destination.resolve() == source.resolve()
                                    or (source.is_dir() and source.resolve() in destination.resolve().parents)):
        raise ValueError("Readiness output must not replace configuration or source data")
    files = DatasetRepository._training_files(source)
    inventory = [{"path": path.relative_to(source).as_posix() if source.is_dir() else path.name,
                  "bytes": path.stat().st_size, "sha256": sha256_file(path)} for path in files]
    model_configs = {name: Path(settings["config"]) for name, settings in values["models"].items()}
    model_configs = {name: path if path.is_absolute() else project_root / path for name, path in model_configs.items()}
    if destination is not None and any(destination.resolve() == path.resolve() for path in model_configs.values()):
        raise ValueError("Readiness output must not replace model configuration")
    model_hashes = {name: sha256_file(path) for name, path in model_configs.items()}
    audits, loading = [], None
    for horizon in values["horizons_hours"]:
        candidates = build_candidate_configs(values, horizon, Path("readiness_only_not_a_training_run"), project_root=project_root)
        # Current experiments share loading policy. Refuse to pretend equivalent
        # numerical populations if future model overrides make those differ.
        policies = {(cfg.values.get("csv_chunk_rows", 100_000), cfg.values.get("memory_limit_mb", 1536),
                     cfg.values.get("numeric_dtype", "float32"), cfg.values.get("energy_source_filter"),
                     require_model_quality_filter(cfg.values)) for _, cfg in candidates}
        if len(policies) != 1:
            raise ValueError("Readiness requires candidates with a shared loading policy")
        if loading is None:
            chunk_rows, memory_limit, dtype, energy, quality = next(iter(policies))
            features = list(dict.fromkeys(column for _, cfg in candidates for column in cfg.values["feature_columns"]))
            _, frame, loading = DatasetRepository(source.parent).load_training_frame(
                source, columns=["timestamp", "plant_id", *features, "generation_mwh"],
                numeric_columns=[*features, "generation_mwh"], equals_filters={"energy_source": energy},
                truthy_filter=quality, policy=DatasetLoadPolicy(chunk_rows, memory_limit, dtype))
        audits.append(audit_forecast_frame(frame, candidates, minimum_coverage=float(values.get("minimum_common_coverage", 0.95))))
    # Prevent report provenance from silently referring to a concurrently edited input.
    try:
        current_files = DatasetRepository._training_files(source)
    except FileNotFoundError as exc:
        raise ValueError("Source dataset or partitions disappeared during readiness auditing") from exc
    if current_files != files:
        raise ValueError("Source partition membership changed during readiness auditing")
    if any(sha256_file(path) != item["sha256"] for path, item in zip(files, inventory)):
        raise ValueError("Source data changed during readiness auditing")
    if sha256_file(config_path) != config_hash or any(sha256_file(path) != model_hashes[name] for name, path in model_configs.items()):
        raise ValueError("Configuration changed during readiness auditing")
    report = {"contract": READINESS_CONTRACT, "source": str(source), "source_files": inventory,
              "experiment_config": str(config_path), "experiment_config_sha256": config_hash,
              "model_config_sha256": model_hashes, "loading": loading.to_dict(), "audits": audits,
              "coverage_gate_passed": all(audit["coverage_gate_passed"] for audit in audits),
              "prediction_or_training_performed": False, "test_used_for_selection": False,
              "limitations": [
                  "Sample counts and input sanity do not establish accuracy or improved model performance.",
                  "Past observations are assumed available at their timestamps; provider publication latency is not simulated.",
                  "Missing input features are retained; measured targets and missing hours are never imputed.",
                  "All-missing or constant Train features are diagnostics, not new exclusion rules.",
                  "Calibration/Test coverage across every candidate is conservative; the benchmark later compares its frozen selected pair.",
                  "Training budgets, runtime memory, convergence, and model/checkpoint execution require separate local model validation."]}
    if destination is not None:
        write_json_atomic(destination, report)
    return report
