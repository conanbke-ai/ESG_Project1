"""실측 예측 실험의 기간·특징·모델별 탐색 예산을 실행 설정으로 변환."""
from __future__ import annotations

from copy import deepcopy
from dataclasses import asdict
import json
from pathlib import Path
import re

from solar_forecast.config_loader import ModelJobConfig, PROJECT_ROOT, load_model_config
from solar_forecast.evaluation.temporal_split import CALENDAR_FIELDS, TemporalSplitConfig, calendar_split_for_execution
from solar_forecast.models.shared.checkpoint_store import BENCHMARK_CHECKPOINT_CONTRACT


EXPERIMENT_CONTRACT = "solar-optimized-experiment.v1"


def load_experiment_config(path: Path, *, project_root: Path = PROJECT_ROOT) -> dict:
    """Validate the executable experiment, including explicitly bounded searches."""
    path = path if path.is_absolute() else project_root / path
    values = json.loads(path.read_text(encoding="utf-8"))
    if values.get("contract") != EXPERIMENT_CONTRACT:
        raise ValueError(f"Experiment requires contract {EXPERIMENT_CONTRACT}")
    split_config = _shared_split(values)
    if values.get("name") == "historical_observed_solar_global_calendar_benchmark" and split_config.split_mode != "calendar":
        raise ValueError("Canonical Solar benchmark requires one global calendar split; per-plant/fraction split is not allowed")
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
    protocol = values.get("selection_protocol")
    if not isinstance(protocol, dict):
        raise ValueError("selection_protocol must be an object")
    if protocol.get("base_candidates") != "validation_only_common_rows":
        raise ValueError(
            "selection_protocol.base_candidates must be validation_only_common_rows"
        )
    if protocol.get("metric") != "pooled_plant_hour_mae":
        raise ValueError(
            "selection_protocol.metric must be pooled_plant_hour_mae"
        )
    if protocol.get("test") != "frozen_models_final_report_only":
        raise ValueError(
            "selection_protocol.test must reserve Test for final reporting"
        )
    future_weather = values.get("future_weather")
    if future_weather is not None:
        if not isinstance(future_weather, dict):
            raise ValueError("future_weather must be an object")
        enabled_horizons = future_weather.get("enabled_horizons", [])
        if any(type(value) is not int or value <= 0 for value in enabled_horizons):
            raise ValueError(
                "future_weather.enabled_horizons must contain positive integers"
            )
        if len(set(enabled_horizons)) != len(enabled_horizons):
            raise ValueError("future_weather.enabled_horizons cannot contain duplicates")
        if enabled_horizons and not future_weather.get("source_template"):
            raise ValueError(
                "future_weather.source_template is required when enabled"
            )
        coverage = float(future_weather.get("minimum_coverage", 0.98))
        if not 0 < coverage <= 1:
            raise ValueError("future_weather.minimum_coverage must be in (0, 1]")
        profiles = future_weather.get(
            "feature_profiles",
            [future_weather.get("feature_profile", "aligned_core")],
        )
        allowed_profiles = {
            "aligned_core",
            "aligned_core_plus_components",
        }
        if (
            not profiles
            or any(str(profile) not in allowed_profiles for profile in profiles)
            or len(set(map(str, profiles))) != len(profiles)
        ):
            raise ValueError(
                "future_weather.feature_profiles must contain unique supported profiles"
            )
    if not values.get("input_dataset"):
        raise ValueError("input_dataset is required")
    return values


def build_candidate_configs(values: dict, horizon: int, run_dir: Path, *, project_root: Path = PROJECT_ROOT) -> list[tuple[str, ModelJobConfig]]:
    """Allow independent features/lookbacks while freezing target and time splits."""
    candidates = []
    shared_split = asdict(_shared_split(values))
    shared_split["purge_gap_hours"] = shared_split.pop("gap_hours")
    for model, settings in values["models"].items():
        config_path = Path(settings["config"])
        config_path = config_path if config_path.is_absolute() else project_root / config_path
        base = load_model_config(config_path)
        if base.model != model:
            raise ValueError(f"Model config does not match {model}")
        lengths = settings.get("sequence_lengths", [1]) if model == "cnn_bilstm" else [1]
        future_weather = values.get("future_weather") or {}
        enabled_horizons = set(future_weather.get("enabled_horizons", []))
        future_profiles = (
            list(
                future_weather.get(
                    "feature_profiles",
                    [future_weather.get("feature_profile", "aligned_core")],
                )
            )
            if horizon in enabled_horizons
            else [None]
        )
        for feature_set in settings["feature_sets"]:
            spec = values["feature_sets"][feature_set]
            features = spec.get("columns", base.values["feature_columns"])
            features = [column for column in features if column not in spec.get("exclude", [])]
            if not features or len(set(features)) != len(features):
                raise ValueError(f"Invalid features for {feature_set}")
            for length in lengths:
                for future_profile in future_profiles:
                    base_candidate_id = (
                        f"{feature_set}_lookback_{length}h"
                        if model == "cnn_bilstm"
                        else feature_set
                    )
                    candidate_id = (
                        f"{base_candidate_id}_future_{future_profile}"
                        if future_profile is not None and len(future_profiles) > 1
                        else base_candidate_id
                    )
                    config = deepcopy(base.values)
                    # Overrides are model-specific; task identity below always wins.
                    config.update(deepcopy(settings.get("training_overrides", {})))
                    # Every candidate uses the same complete split specification,
                    # including unset dates; model overrides cannot alter its calendar.
                    config.update(shared_split)
                    if values.get("smoke_split_override"):
                        config["smoke_split_override"] = deepcopy(values["smoke_split_override"])
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
                        "checkpoint_identity_contract": BENCHMARK_CHECKPOINT_CONTRACT,
                    })
                    if horizon in enabled_horizons:
                        source_template = str(future_weather["source_template"])
                        source_value = source_template.format(horizon=horizon)
                        config["future_weather"] = {
                            "enabled": True,
                            "source": str(_resolve_path(source_value, project_root)),
                            "minimum_coverage": float(
                                future_weather.get("minimum_coverage", 0.98)
                            ),
                            "source_contract": str(
                                future_weather.get(
                                    "source_contract",
                                    "solar-future-weather-forecast.v2",
                                )
                            ),
                            "feature_profile": str(
                                future_profile or "aligned_core"
                            ),
                        }
                    else:
                        config["future_weather"] = {"enabled": False}
    
                    if model == "cnn_bilstm":
                        from solar_forecast.models.cnn_bilstm.input_preprocessing import INPUT_PREPROCESSING_CONTRACT
    
                        config["sequence_length"] = length
                        config["input_preprocessing_contract"] = INPUT_PREPROCESSING_CONTRACT
                    config["purge_gap_hours"] = max(int(config.get("purge_gap_hours", 168)), horizon)
                    TemporalSplitConfig.from_mapping(config)
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
        planned_candidates = []
        for name, cfg in candidates:
            future = cfg.values.get("future_weather") or {}
            future_enabled = bool(future.get("enabled", False))
            future_profile = (
                str(future.get("feature_profile", "aligned_core"))
                if future_enabled
                else None
            )
            if future_enabled:
                from solar_forecast.features.future_weather import (
                    future_weather_feature_columns,
                )

                future_features = list(
                    future_weather_feature_columns(str(future_profile))
                )
            else:
                future_features = []
            planned_candidates.append(
                {
                    "model": cfg.model,
                    "candidate_id": name,
                    "historical_features": cfg.values["feature_columns"],
                    "future_weather_enabled": future_enabled,
                    "future_weather_profile": future_profile,
                    "future_weather_features": future_features,
                    "sequence_length": cfg.values.get("sequence_length"),
                    "target_transform": cfg.values.get(
                        "target_transform", "identity"
                    ),
                    "max_trials": cfg.values["optimizer"]["max_trials"],
                    "timeout_seconds": cfg.values["optimizer"]["timeout_seconds"],
                }
            )
        tasks.append(
            {
                "horizon_hours": horizon,
                "candidates": planned_candidates,
            }
        )
    split = _shared_split(values)
    return {"contract": EXPERIMENT_CONTRACT, "input_dataset": str(_resolve_path(values["input_dataset"], project_root)),
            "split": {"mode": split.split_mode, **asdict(split)},
            "smoke_split_override": values.get("smoke_split_override"),
            "prediction_task": "historical_forecast", "tasks": tasks,
            "selection_protocol": "validation_base_search_then_calibration_gate_fit_and_selection_then_test_report",
            "optimization_scope": values.get("optimization_scope", "configured_search_budget")}


def configure_smoke_experiment(values: dict) -> dict:
    """Create the bounded wiring plan without pretending to test calendar holdouts."""
    smoke_values = deepcopy(values)
    fields, override = calendar_split_for_execution(smoke_values.get("split", {}), smoke=True)
    smoke_values["horizons_hours"] = [1]
    smoke_values["selection_gap_hours"] = 0
    smoke_values.setdefault("split", {}).update(fields, purge_gap_hours=0)
    smoke_values["optimization_scope"] = "smoke_wiring_only"
    if override is not None:
        smoke_values["smoke_split_override"] = override
    for model, settings in smoke_values["models"].items():
        settings["feature_sets"] = settings["feature_sets"][:1]
        if model == "cnn_bilstm":
            settings["sequence_lengths"] = [24]
    return smoke_values


def _shared_split(values: dict) -> TemporalSplitConfig:
    split = values.get("split", {})
    if not isinstance(split, dict):
        raise ValueError("split must be an object")
    supported = {*CALENDAR_FIELDS, "validation_fraction", "calibration_fraction", "test_fraction", "purge_gap_hours"}
    if unknown := set(split) - supported:
        raise ValueError(f"Unknown split fields: {sorted(unknown)}")
    return TemporalSplitConfig.from_mapping({"purge_gap_hours": 168, **split})


def _resolve_path(value: str, project_root: Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else project_root / path
