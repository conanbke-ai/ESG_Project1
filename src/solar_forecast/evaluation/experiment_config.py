"""실측 예측 실험의 기간·특징·모델별 탐색 예산을 실행 설정으로 변환."""
from __future__ import annotations

from copy import deepcopy
import json
from pathlib import Path
import re

from solar_forecast.config_loader import ModelJobConfig, PROJECT_ROOT, load_model_config


EXPERIMENT_CONTRACT = "solar-optimized-experiment.v1"


def load_experiment_config(path: Path, *, project_root: Path = PROJECT_ROOT) -> dict:
    """Validate the executable experiment, including explicitly bounded searches."""
    path = path if path.is_absolute() else project_root / path
    values = json.loads(path.read_text(encoding="utf-8"))
    if values.get("contract") != EXPERIMENT_CONTRACT:
        raise ValueError(f"Experiment requires contract {EXPERIMENT_CONTRACT}")
    horizons = values.get("horizons_hours", [])
    if not horizons or any(type(h) is not int or h <= 0 for h in horizons) or len(set(horizons)) != len(horizons):
        raise ValueError("horizons_hours must contain distinct positive integers")
    if not 0 < float(values.get("minimum_common_coverage", 0.95)) <= 1:
        raise ValueError("minimum_common_coverage must be in (0, 1]")
    if not 0 <= float(values.get("minimum_relative_improvement", 0)) < 1:
        raise ValueError("minimum_relative_improvement must be in [0, 1)")
    if set(values.get("models", {})) != {"xgboost", "cnn_bilstm"}:
        raise ValueError("Both xgboost and cnn_bilstm model searches are required")
    for name, settings in values["models"].items():
        if not settings.get("feature_sets") or not settings.get("config"):
            raise ValueError(f"{name}: model config and feature_sets are required")
        if any(key not in values.get("feature_sets", {}) for key in settings["feature_sets"]):
            raise ValueError(f"{name}: unknown feature set")
        if any(not re.fullmatch(r"[a-z][a-z0-9_]*", key) for key in settings["feature_sets"]):
            raise ValueError("Feature set IDs must be snake_case")
        lengths = settings.get("sequence_lengths", [1])
        if not lengths or any(type(n) is not int or n <= 0 for n in lengths):
            raise ValueError(f"{name}: sequence lengths must be positive integers")
        if len(set(lengths)) != len(lengths) or len(set(settings["feature_sets"])) != len(settings["feature_sets"]):
            raise ValueError(f"{name}: duplicate candidates")
        for field in ("max_trials_per_candidate", "timeout_seconds_per_candidate"):
            if type(settings.get(field)) is not int or settings[field] < 1:
                raise ValueError(f"{name}: {field} must be a positive integer")
    if not values.get("input_dataset"):
        raise ValueError("input_dataset is required")
    return values


def build_candidate_configs(values: dict, horizon: int, run_dir: Path, *, project_root: Path = PROJECT_ROOT) -> list[tuple[str, ModelJobConfig]]:
    """Allow independent features/lookbacks while freezing target and time splits."""
    candidates = []
    for model, settings in values["models"].items():
        config_path = Path(settings["config"])
        config_path = config_path if config_path.is_absolute() else project_root / config_path
        base = load_model_config(config_path)
        if base.model != model:
            raise ValueError(f"Model config does not match {model}")
        lengths = settings.get("sequence_lengths", [1]) if model == "cnn_bilstm" else [1]
        for feature_set in settings["feature_sets"]:
            spec = values["feature_sets"][feature_set]
            features = spec.get("columns", base.values["feature_columns"])
            features = [column for column in features if column not in spec.get("exclude", [])]
            if not features or len(set(features)) != len(features):
                raise ValueError(f"Invalid features for {feature_set}")
            for length in lengths:
                candidate_id = f"{feature_set}_lookback_{length}h" if model == "cnn_bilstm" else feature_set
                config = deepcopy(base.values)
                # Overrides are model-specific; task identity below always wins.
                config.update(deepcopy(settings.get("training_overrides", {})))
                config.update(deepcopy(values.get("split", {})))
                config.update({
                    "input_dataset": str(_resolve_path(values["input_dataset"], project_root)),
                    "target_column": "generation_mwh", "energy_source_filter": "solar",
                    "quality_filter_column": "quality_train_eligible",
                    "prediction_task": "historical_forecast", "forecast_horizon_hours": horizon,
                    "evaluation_protocol": "historical_observation_rolling_origin",
                    "feature_columns": features, "feature_contract": f"historical_origin_v1:{feature_set}",
                    "seed": int(values.get("seed", 42)),
                    "output_root": str(run_dir / "candidates" / f"horizon_{horizon}h" / model / candidate_id),
                    "benchmark_candidate_id": candidate_id,
                })
                if model == "cnn_bilstm":
                    config["sequence_length"] = length
                config["purge_gap_hours"] = max(int(config.get("purge_gap_hours", 168)), horizon)
                config["optimizer"] = {
                    **config.get("optimizer", {}),
                    **deepcopy(settings.get("optimizer_overrides", {})),
                    "enabled": bool(settings.get("optimize", True)),
                    "max_trials": settings["max_trials_per_candidate"],
                    "timeout_seconds": settings["timeout_seconds_per_candidate"],
                    "study_name": f"historical_{model}_{horizon}h_{candidate_id}",
                }
                candidates.append((candidate_id, ModelJobConfig(model, "historical_optimized", config, config_path)))
    return candidates


def experiment_plan(values: dict, *, project_root: Path = PROJECT_ROOT) -> dict:
    """Describe actual candidates without loading data or starting training."""
    tasks = []
    for horizon in values["horizons_hours"]:
        candidates = build_candidate_configs(values, horizon, Path(values.get("output_root", "artifacts/benchmarks")), project_root=project_root)
        tasks.append({"horizon_hours": horizon, "candidates": [
            {"model": cfg.model, "candidate_id": name, "features": cfg.values["feature_columns"],
             "sequence_length": cfg.values.get("sequence_length"),
             "max_trials": cfg.values["optimizer"]["max_trials"],
             "timeout_seconds": cfg.values["optimizer"]["timeout_seconds"]}
            for name, cfg in candidates
        ]})
    return {"contract": EXPERIMENT_CONTRACT, "input_dataset": str(_resolve_path(values["input_dataset"], project_root)),
            "prediction_task": "historical_forecast", "tasks": tasks,
            "selection_protocol": "validation_base_search_then_calibration_gate_fit_and_selection_then_test_report",
            "optimization_scope": values.get("optimization_scope", "configured_search_budget")}


def _resolve_path(value: str, project_root: Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else project_root / path
