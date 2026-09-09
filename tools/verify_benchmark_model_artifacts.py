"""Replay selected benchmark checkpoints on observed Test rows without training.

This checks the newly saved models and their stored preprocessing. It does not
reproduce the project's earlier notebook/checkpoint performance.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import sys

import numpy as np
import pandas as pd


KEYS = ["plant_id", "timestamp", "forecast_origin", "horizon_hours"]


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _resolve(value: str, run_dir: Path) -> Path:
    path = Path(value)
    if path.exists():
        return path
    if not path.is_absolute():
        path = run_dir / path
    elif run_dir.name in path.parts:
        # Permit replay after moving the complete benchmark artifact folder.
        suffix = path.parts[path.parts.index(run_dir.name) + 1:]
        path = run_dir.joinpath(*suffix)
    if not path.exists():
        raise FileNotFoundError(f"Stored model artifact is missing: {value}")
    return path


def _load_observations(dataset: Path, values: dict, *, smoke: bool) -> pd.DataFrame:
    from solar_forecast.datasets.numeric_preprocessor import NumericPreprocessor, require_model_quality_filter
    from solar_forecast.datasets.repository import DatasetLoadPolicy, DatasetRepository
    from solar_forecast.evaluation.forecast_samples import validate_observation_frame

    target = values["target_column"]
    features = values["feature_columns"]
    passthrough = ["plant_id", "timestamp", "plant", "region"]
    _, frame, _ = DatasetRepository(dataset.parent).load_training_frame(
        dataset, columns=[*passthrough, *features, target],
        numeric_columns=[*features, target],
        equals_filters={"energy_source": str(values["energy_source_filter"])},
        truthy_filter=require_model_quality_filter(values),
        row_limit=10000 if smoke else None,
        policy=DatasetLoadPolicy(
            numeric_dtype=str(values.get("numeric_dtype", "float32")),
            chunk_rows=int(values.get("csv_chunk_rows", 100000)),
            memory_limit_mb=int(values.get("memory_limit_mb", 1536)),
        ),
    )
    # This stateless step only selects/casts columns. No median or scaler is fit.
    frame = NumericPreprocessor(fill_missing=False).transform(
        frame, target, features, passthrough_columns=passthrough,
    ).frame
    if smoke:
        first = frame["plant_id"].iloc[0]
        frame = frame.loc[frame["plant_id"].eq(first)].head(512)
    frame = validate_observation_frame(frame, entity_column="plant_id", timestamp_column="timestamp")
    frame["plant_id"] = frame["plant_id"].astype(str)
    return frame


def _test_start(split: dict) -> pd.Timestamp:
    boundaries = split["boundaries"]
    return pd.Timestamp(boundaries["calibration_end"]) + pd.Timedelta(hours=int(boundaries["gap_hours"]))


def _indexed(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.copy()
    result["plant_id"] = result["plant_id"].astype(str)
    for column in ("timestamp", "forecast_origin"):
        result[column] = pd.to_datetime(result[column], errors="raise")
    if result.empty or result[KEYS].isna().any().any() or result.duplicated(KEYS).any():
        raise ValueError("Stored Test predictions contain empty, null, or duplicate forecast keys")
    return result.set_index(KEYS)


def _match_truth(expected: pd.DataFrame, exported: pd.DataFrame, *, target_dtype: str) -> dict:
    if len(expected) != len(exported) or set(expected.index) != set(exported.index):
        raise ValueError("Reconstructed Test forecast keys differ from the saved prediction artifact")
    expected = expected.loc[exported.index]
    report = {}
    for column in ("y_true", "persistence_pred"):
        # Training CSV writers may shorten float32 decimal text. Compare the
        # underlying model-input dtype, not the display precision of that text.
        actual = exported[column].to_numpy(dtype=target_dtype)
        reference = expected[column].to_numpy(dtype=target_dtype)
        if not np.isfinite(actual).all() or not np.array_equal(actual, reference):
            raise ValueError(f"Observed data does not reproduce saved {column}")
        report[f"{column}_max_abs_delta_at_input_precision"] = float(
            np.max(np.abs(actual.astype(float) - reference.astype(float)))
        )
    return report


def _xgboost_replay(frame: pd.DataFrame, values: dict, details: dict, exported: pd.DataFrame, run_dir: Path):
    from xgboost import XGBRegressor
    from solar_forecast.evaluation.forecast_samples import build_forecast_samples

    model_path = _resolve(details["model_path"], run_dir)
    preprocessing_path = _resolve(details["preprocessing_path"], run_dir)
    preprocessing = _json(preprocessing_path)
    if preprocessing["strategy"] != "xgboost_native_nan_learned_default_direction":
        raise ValueError("Unsupported stored XGBoost preprocessing")
    samples = build_forecast_samples(
        frame, values["feature_columns"], values["target_column"], values["forecast_horizon_hours"],
    )
    samples = samples.loc[samples["timestamp"].gt(_test_start(preprocessing["temporal_split"]))]
    expected = _indexed(samples.rename(columns={values["target_column"]: "y_true"}))
    truth = _match_truth(expected, exported, target_dtype=str(values.get("numeric_dtype", "float32")))
    model = XGBRegressor(n_jobs=1)
    model.load_model(model_path)
    if model.get_booster().feature_names != list(values["feature_columns"]):
        raise ValueError("Saved XGBoost feature order differs from resolved configuration")
    predicted = model.predict(expected.loc[exported.index, values["feature_columns"]])
    return predicted, {
        **truth, "model_path": str(model_path), "model_sha256": _sha256(model_path),
        "preprocessing_sha256": _sha256(preprocessing_path),
        "preprocessing_source": "saved_native_nan_contract",
    }


def _cnn_replay(frame: pd.DataFrame, values: dict, details: dict, exported: pd.DataFrame, run_dir: Path):
    import torch
    from solar_forecast.evaluation.forecast_samples import forecast_window_positions
    from solar_forecast.models.cnn_bilstm.network import CnnBiLstmNetworkConfig, build_cnn_bilstm_network

    model_path = _resolve(details["checkpoint_path"], run_dir)
    checkpoint = torch.load(model_path, map_location="cpu", weights_only=True)
    features = checkpoint["feature_columns"]
    if features != list(values["feature_columns"]):
        raise ValueError("Stored CNN feature order differs from resolved configuration")
    preprocessing = checkpoint["preprocessing"]
    if preprocessing["strategy"] != "training_split_median_with_missing_indicators":
        raise ValueError("Unsupported stored CNN preprocessing")
    split = preprocessing["temporal_split"]
    length = int(split["sequence_length"])
    medians = np.array([preprocessing["feature_medians"][name] for name in features], dtype=np.float32)
    if not np.isfinite(medians).all():
        raise ValueError("Stored CNN medians must be finite")
    series = {}
    expected_frames = []
    origin_by_key = {}
    horizon = int(values["forecast_horizon_hours"])
    for plant, group in frame.groupby("plant_id", sort=True, observed=True):
        group = group.sort_values("timestamp", kind="stable")
        numeric = group[features].to_numpy(dtype=np.float32, copy=True)
        missing = ~np.isfinite(numeric)
        missing_rows, missing_columns = np.where(missing)
        numeric[missing_rows, missing_columns] = medians[missing_columns]
        if preprocessing["append_missing_indicators"]:
            numeric = np.concatenate([numeric, missing.astype(np.float32)], axis=1)
        if numeric.shape[1] != checkpoint["config"]["n_features"]:
            raise ValueError("Stored CNN input width differs from preprocessing")
        times = pd.DatetimeIndex(group["timestamp"])
        target_values = group[values["target_column"]].to_numpy(dtype=np.float32)
        targets, origins = forecast_window_positions(times, horizon_hours=horizon, sequence_length=length)
        keep = (times[targets] > _test_start(split)) & np.isfinite(target_values[targets]) & np.isfinite(target_values[origins])
        targets, origins = targets[keep], origins[keep]
        series[str(plant)] = numeric
        context = pd.DataFrame({
            "plant_id": str(plant), "timestamp": times[targets],
            "forecast_origin": times[origins], "horizon_hours": horizon,
            "y_true": target_values[targets], "persistence_pred": target_values[origins],
        })
        expected_frames.append(context)
        for target, origin in zip(targets, origins):
            origin_by_key[(str(plant), times[target], times[origin], horizon)] = int(origin)
    expected = _indexed(pd.concat(expected_frames, ignore_index=True))
    truth = _match_truth(expected, exported, target_dtype="float32")
    model = build_cnn_bilstm_network(CnnBiLstmNetworkConfig(**checkpoint["config"]), device=torch.device("cpu"))
    model.load_state_dict(checkpoint["model_state"], strict=True)
    model.eval()
    predictions = []
    keys = list(exported.index)
    batch_size = int(values.get("batch_size", 128))
    with torch.no_grad():
        for start in range(0, len(keys), batch_size):
            windows = []
            for key in keys[start:start + batch_size]:
                origin = origin_by_key[key]
                windows.append(series[key[0]][origin - length + 1:origin + 1])
            batch = torch.from_numpy(np.stack(windows)).float()
            predictions.append(model(batch).cpu().numpy().reshape(-1))
    return np.concatenate(predictions), {
        **truth, "model_path": str(model_path), "model_sha256": _sha256(model_path),
        "preprocessing_sha256": hashlib.sha256(json.dumps(preprocessing, sort_keys=True).encode()).hexdigest(),
        "preprocessing_source": "checkpoint_stored_training_medians_and_missing_indicators",
        "sequence_length": length, "batch_size": batch_size,
    }


def verify_selected_artifacts(run_dir: Path, dataset: Path, *, prediction_atol: float = 1e-6) -> dict:
    """Reload each selected base model once; return evidence without fitting anything."""

    run_dir, dataset = Path(run_dir).resolve(), Path(dataset).resolve()
    manifest = _json(run_dir / "manifest.json")
    digest = _sha256(dataset)
    if manifest.get("status") != "completed" or manifest["provenance"]["dataset_fingerprint"] != digest:
        raise ValueError("Replay requires a completed benchmark and its exact observed dataset")
    results = []
    for task in manifest["tasks"]:
        for name, chosen in task["optimization"].items():
            result = {"model": name, "horizon_hours": task["horizon_hours"], "status": "failed"}
            try:
                candidate_dir = _resolve(chosen["run_dir"], run_dir)
                details = _json(candidate_dir / "manifest.json")["details"]
                config_path = _resolve(chosen["resolved_config"], run_dir)
                values = _json(config_path)
                if values.get("prediction_task") != "historical_forecast" or values["forecast_horizon_hours"] != task["horizon_hours"]:
                    raise ValueError("Replay only supports the declared historical forecast task")
                contract = details["evaluation_contract"]
                if contract["dataset_fingerprint"] != digest:
                    raise ValueError("Selected base model uses a different observed dataset")
                prediction_path = _resolve(details["test_predictions"], run_dir)
                exported = _indexed(pd.read_csv(prediction_path, dtype={"plant_id": str}))
                frame = _load_observations(dataset, values, smoke=manifest.get("execution_mode") == "smoke")
                replay = {"xgboost": _xgboost_replay, "cnn_bilstm": _cnn_replay}[name]
                predicted, evidence = replay(frame, values, details, exported, run_dir)
                saved = exported["y_pred"].to_numpy(dtype=float)
                delta = np.abs(np.asarray(predicted, dtype=float) - saved)
                passed = np.isfinite(delta).all() and np.all(delta <= prediction_atol)
                result.update(evidence)
                result.update({
                    "status": "passed" if passed else "failed", "rows": len(exported),
                    "prediction_max_abs_delta": float(delta.max()),
                    "prediction_mean_abs_delta": float(delta.mean()),
                    "prediction_atol": prediction_atol,
                    "prediction_artifact_sha256": _sha256(prediction_path),
                    "resolved_config_sha256": _sha256(config_path),
                })
            except Exception as exc:
                result["error"] = f"{type(exc).__name__}: {exc}"
            results.append(result)
    return {
        "contract": "solar-benchmark-stored-model-replay.v1",
        "status": "passed" if results and all(item["status"] == "passed" for item in results) else "failed",
        "dataset_sha256": digest, "models": results,
        "training_performed": False, "preprocessors_refit": False,
        "historical_checkpoint_performance_verified": False,
        "scope": "selected_new_benchmark_artifacts_on_all_their_observed_Test_rows",
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("run_dir", type=Path)
    parser.add_argument("dataset", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
    report = verify_selected_artifacts(args.run_dir, args.dataset)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    raise SystemExit(0 if report["status"] == "passed" else 1)


if __name__ == "__main__":
    main()
