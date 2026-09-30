"""완료된 실측 데이터 벤치마크의 선택 근거와 최종 평가를 화면에 투영."""
from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


MODEL_LABELS = {
    "xgboost": "XGBoost",
    "cnn_bilstm": "CNN-BiLSTM",
    "hybrid": "Hybrid",
    "persistence": "직전 관측 유지 기준선",
}
PREDICTION_COLUMNS = {
    "xgboost": "xgb_pred",
    "cnn_bilstm": "cnn_pred",
    "hybrid": "hybrid_pred",
    "persistence": "persistence_pred",
}


class BenchmarkAnalyticsService:
    """Keep benchmark selection evidence separate from legacy anomaly policies."""

    def __init__(self, project_root: Path):
        self.project_root = Path(project_root).resolve()
        self.artifact_root = self.project_root / "artifacts/benchmarks"

    def build(self) -> dict[str, Any]:
        excluded = {"non_full_or_incomplete": 0, "invalid_artifacts": 0}
        manifests = sorted(
            self.artifact_root.glob("*/manifest.json"),
            key=lambda path: (path.stat().st_mtime_ns, path.parent.name),
            reverse=True,
        )
        for path in manifests:
            try:
                manifest = _read_json(path)
                if (
                    manifest.get("contract") != "solar-optimized-benchmark.v1"
                    or manifest.get("status") != "completed"
                    or manifest.get("execution_mode") != "full"
                ):
                    excluded["non_full_or_incomplete"] += 1
                    continue
                tasks = manifest.get("tasks", [])
                if not tasks:
                    raise ValueError("Benchmark has no evaluation tasks")
                projected = [self._task(path.parent, manifest, task) for task in tasks]
                horizons = [task["horizon_hours"] for task in projected]
                if len(horizons) != len(set(horizons)):
                    raise ValueError("Duplicate benchmark horizons")
                return {
                    "status": "ready",
                    "run_id": path.parent.name,
                    "execution_mode": "full",
                    "message": "실측 데이터의 과거 Test 구간 평가입니다. 모델 선택은 Test 이전 구간에서 완료했습니다.",
                    "tasks": sorted(projected, key=lambda task: task["horizon_hours"]),
                    "excluded_runs": excluded,
                }
            except (OSError, ValueError, KeyError, TypeError, OverflowError):
                excluded["invalid_artifacts"] += 1
        return {
            "status": "empty",
            "message": "완료된 정식 벤치마크 결과가 없습니다. 실행 점검용 결과는 성능 비교에 포함하지 않습니다.",
            "tasks": [],
            "excluded_runs": excluded,
        }

    def _task(
        self, run_dir: Path, manifest: dict[str, Any], task: dict[str, Any]
    ) -> dict[str, Any]:
        selection_path = _contained_path(run_dir, task["selection_path"])
        selection = _read_json(selection_path)
        if selection.get("contract") != "solar-benchmark-selection.v1" or selection.get("status") != "completed":
            raise ValueError("Selection is incomplete")
        horizon = int(task["horizon_hours"])
        if horizon <= 0 or horizon != task["horizon_hours"]:
            raise ValueError("Invalid horizon")
        evaluation = selection["evaluation_contract"]
        if evaluation.get("horizon_hours") != horizon or evaluation.get("task") != "historical_forecast" or evaluation.get("information_set") != "observed_measurements_through_forecast_origin":
            raise ValueError("Selection horizon differs from task")
        fingerprint = manifest.get("provenance", {}).get("dataset_fingerprint")
        if not fingerprint or selection.get("provenance", {}).get("dataset_fingerprint") != fingerprint:
            raise ValueError("Dataset provenance does not match")
        selected = selection["selected_model"]
        if selected not in {"xgboost", "cnn_bilstm", "hybrid"}:
            raise ValueError("Invalid selected model")
        prediction_path = _contained_path(selection_path.parent, selection["files"]["test_predictions"])
        expected_digest = selection.get("files_sha256", {}).get("test_predictions")
        if not expected_digest or sha256(prediction_path.read_bytes()).hexdigest() != expected_digest:
            raise ValueError("Test prediction fingerprint differs")
        frame = pd.read_csv(prediction_path, dtype={"plant_id": str})
        _validate_predictions(frame, horizon, selected, selection["periods"])
        models = []
        for model_id, label in MODEL_LABELS.items():
            selection_metrics = selection["selection_metrics"][model_id]
            test_metrics = selection["test_metrics"][model_id]
            _validate_metric_tree(selection_metrics)
            _validate_metric_tree(test_metrics)
            _verify_test_metrics(frame, PREDICTION_COLUMNS[model_id], test_metrics)
            models.append({
                "id": model_id,
                "label": label,
                "selected": model_id == selected,
                "selection_metrics": selection_metrics,
                "test_metrics": test_metrics,
            })
        rule = selection["selection_rule"]
        margin = float(rule["minimum_relative_improvement"])
        if rule.get("test_used_for_selection") is not False or not 0 <= margin < 1:
            raise ValueError("Invalid benchmark selection rule")
        scores = {model: selection["selection_metrics"][model]["pooled"]["mae"] for model in MODEL_LABELS}
        best_base = min(("xgboost", "cnn_bilstm"), key=lambda model: scores[model])
        expected_selected = "hybrid" if scores["hybrid"] < scores[best_base] * (1 - margin) else best_base
        if selected != expected_selected:
            raise ValueError("Selected model differs from selection evidence")
        return {
            "horizon_hours": horizon,
            "selected_model": selected,
            "selected_label": MODEL_LABELS[selected],
            "reason": (
                "선택 구간에서 Hybrid의 MAE가 가장 좋은 단일 모델보다 낮고, 요구 개선율을 충족해 채택했습니다."
                if selected == "hybrid" else
                "선택 구간에서 Hybrid가 요구 개선율을 충족하지 않아 MAE가 가장 낮은 단일 모델을 채택했습니다."
            ),
            "models": models,
            "periods": selection["periods"],
            "persistence_comparison": selection.get("persistence_comparison", {}),
            "alignment": {
                split: {key: values[key] for key in ("common_rows", "union_rows", "common_fraction") if key in values}
                for split, values in task.get("coverage", {}).items()
                if split in {"validation", "calibration", "test"}
            },
            "provenance": {
                "dataset_fingerprint": fingerprint,
                "optimization_scope": manifest.get("provenance", {}).get("optimization_scope", ""),
            },
            "optimization": {
                model: {
                    key: value for key, value in settings.items()
                    if key in {"candidate_id", "sequence_length", "feature_columns", "validation_mae", "parameters"}
                }
                for model, settings in task.get("optimization", {}).items()
                if model in {"xgboost", "cnn_bilstm"}
            },
            "coverage": {
                "test_rows": len(frame),
                "plants": int(frame["plant_id"].nunique()),
                "regions": int(frame["region"].nunique()),
                "start": frame["timestamp"].min().isoformat(),
                "end": frame["timestamp"].max().isoformat(),
            },
            "series": _recent_series(frame),
        }


def _read_json(path: Path) -> dict[str, Any]:
    result = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(result, dict):
        raise ValueError("Expected an artifact object")
    return result


def _contained_path(root: Path, value: str) -> Path:
    path = Path(value)
    resolved = (path if path.is_absolute() else root / path).resolve()
    if not resolved.is_relative_to(root.resolve()):
        raise ValueError("Artifact is outside its benchmark directory")
    return resolved


def _validate_predictions(
    frame: pd.DataFrame, horizon: int, selected: str, periods: dict[str, Any]
) -> None:
    required = {"timestamp", "forecast_origin", "horizon_hours", "plant_id", "plant", "region", "y_true", "selected_pred", *PREDICTION_COLUMNS.values()}
    if frame.empty or not required.issubset(frame.columns) or frame[list(required)].isna().any().any():
        raise ValueError("Incomplete merged Test predictions")
    for column in ("timestamp", "forecast_origin"):
        frame[column] = pd.to_datetime(frame[column], errors="raise", utc=True)
    if frame.duplicated(["plant_id", "timestamp", "forecast_origin"]).any():
        raise ValueError("Duplicate Test predictions")
    if not (pd.to_numeric(frame["horizon_hours"]) == horizon).all():
        raise ValueError("Mixed Test horizons")
    if not ((frame["timestamp"] - frame["forecast_origin"]) == pd.Timedelta(hours=horizon)).all():
        raise ValueError("Incorrect forecast origin")
    numeric = frame[["y_true", "selected_pred", *PREDICTION_COLUMNS.values()]].to_numpy(float)
    if not np.isfinite(numeric).all():
        raise ValueError("Non-finite Test predictions")
    if not np.array_equal(frame["selected_pred"].to_numpy(float), frame[PREDICTION_COLUMNS[selected]].to_numpy(float)):
        raise ValueError("Selected predictions differ from selected model")
    test = periods["test"]
    if int(test["rows"]) != len(frame):
        raise ValueError("Test sample count differs")
    if pd.to_datetime(test["start"], utc=True) != frame["timestamp"].min() or pd.to_datetime(test["end"], utc=True) != frame["timestamp"].max():
        raise ValueError("Test period differs")
    for left, right in (("gate_fit", "selection"), ("selection", "test")):
        if pd.to_datetime(periods[left]["end"], utc=True) >= pd.to_datetime(periods[right]["start"], utc=True):
            raise ValueError("Benchmark selection periods overlap")


def _validate_metric_tree(metrics: dict[str, Any]) -> None:
    for item in [metrics["pooled"], metrics["national"], *metrics["plant"], *metrics["region"]]:
        if int(item["n_samples"]) < 1:
            raise ValueError("Empty metric population")
        for key in ("mae", "rmse"):
            if not np.isfinite(float(item[key])) or float(item[key]) < 0:
                raise ValueError("Invalid metric value")
        if item.get("r2") is not None and not np.isfinite(float(item["r2"])):
            raise ValueError("Invalid R squared")


def _verify_metric_values(actual: pd.DataFrame, column: str, expected: dict[str, Any]) -> None:
    error = actual["y_true"].to_numpy(float) - actual[column].to_numpy(float)
    if len(actual) != int(expected["n_samples"]):
        raise ValueError("Metric sample count differs from predictions")
    computed = {"mae": np.mean(np.abs(error)), "rmse": np.sqrt(np.mean(error ** 2))}
    if any(not np.isclose(computed[key], float(expected[key]), rtol=1e-8, atol=1e-10) for key in computed):
        raise ValueError("Metric values differ from predictions")
    actual_values = actual["y_true"].to_numpy(float)
    variance = float(np.square(actual_values - actual_values.mean()).sum())
    expected_r2 = expected.get("r2")
    if len(actual_values) <= 1 or variance == 0:
        if expected_r2 is not None:
            raise ValueError("R squared is undefined for constant targets")
    elif expected_r2 is None or not np.isclose(1 - np.square(error).sum() / variance, float(expected_r2), rtol=1e-8, atol=1e-10):
        raise ValueError("R squared differs from predictions")


def _verify_test_metrics(frame: pd.DataFrame, column: str, metrics: dict[str, Any]) -> None:
    _verify_metric_values(frame, column, metrics["pooled"])
    grouped = frame.groupby("timestamp")[["y_true", column]].sum()
    _verify_metric_values(grouped, column, metrics["national"])
    plants = {str(row["plant_id"]): row for row in metrics["plant"]}
    regions = {str(row["region"]): row for row in metrics["region"]}
    if set(plants) != set(frame["plant_id"]) or set(regions) != set(frame["region"]):
        raise ValueError("Metric coverage differs from predictions")
    for plant, group in frame.groupby("plant_id"):
        _verify_metric_values(group, column, plants[str(plant)])
    for region, group in frame.groupby("region"):
        summed = group.groupby("timestamp")[["y_true", column]].sum()
        _verify_metric_values(summed, column, regions[str(region)])


def _recent_series(frame: pd.DataFrame) -> list[dict[str, Any]]:
    result = []
    for _, group in frame.groupby("plant_id", sort=True):
        group = group.sort_values("timestamp")
        blocks = group["timestamp"].diff().ne(pd.Timedelta(hours=1)).cumsum()
        recent = group.loc[blocks == blocks.iloc[-1]].tail(168)
        for row in recent.to_dict("records"):
            result.append({
                "timestamp": row["timestamp"].isoformat(),
                "forecast_origin": row["forecast_origin"].isoformat(),
                "plant_id": str(row["plant_id"]),
                "plant": str(row["plant"]),
                "region": str(row["region"]),
                "y_true": float(row["y_true"]),
                "predictions": {model: float(row[column]) for model, column in PREDICTION_COLUMNS.items()},
            })
    return result


__all__ = ["BenchmarkAnalyticsService"]
