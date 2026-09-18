"""시간 순서로 하이브리드 가중치 학습·모델 채택·최종 평가를 분리."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from solar_forecast.infrastructure.artifact_store import (
    replace_file_atomic,
    sha256_file,
    write_json_atomic,
)
from solar_forecast.models.hybrid.dynamic_gate import ExplainableDynamicGate


SELECTION_CONTRACT = "solar-benchmark-selection.v1"
_PREDICTIONS = {
    "xgboost": "xgb_pred",
    "cnn_bilstm": "cnn_pred",
    "hybrid": "hybrid_pred",
    "persistence": "persistence_pred",
}
_REQUIRED = {
    "timestamp", "forecast_origin", "horizon_hours", "plant_id", "plant",
    "region", "y_true", "xgb_pred", "cnn_pred", "persistence_pred",
}


def _validate_frame(frame: pd.DataFrame, name: str, horizon: int) -> pd.DataFrame:
    missing = _REQUIRED - set(frame.columns)
    if missing:
        raise ValueError(f"{name}: missing prediction columns: {sorted(missing)}")
    if frame.empty or frame.columns.duplicated().any():
        raise ValueError(f"{name}: predictions must be nonempty with unique columns")
    result = frame.copy()
    for column in ("timestamp", "forecast_origin"):
        result[column] = pd.to_datetime(result[column], errors="raise")
        if result[column].isna().any():
            raise ValueError(f"{name}: {column} contains missing times")
        if not pd.api.types.is_datetime64_any_dtype(result[column]):
            raise ValueError(f"{name}: {column} mixes timezone conventions")
    if str(result.timestamp.dt.tz) != str(result.forecast_origin.dt.tz):
        raise ValueError(f"{name}: target and origin timezone conventions differ")
    for column in ("plant_id", "plant", "region"):
        if result[column].isna().any() or result[column].astype(str).str.strip().eq("").any():
            raise ValueError(f"{name}: {column} contains missing identities")
        result[column] = result[column].astype(str)
    for column in ("horizon_hours", "y_true", "xgb_pred", "cnn_pred", "persistence_pred"):
        result[column] = pd.to_numeric(result[column], errors="raise")
        if not np.isfinite(result[column].to_numpy(dtype=float)).all():
            raise ValueError(f"{name}: {column} contains nonfinite values")
    if not result.horizon_hours.eq(horizon).all():
        raise ValueError(f"{name}: horizons differ from the evaluation contract")
    if not (result.timestamp - result.forecast_origin).eq(pd.Timedelta(hours=horizon)).all():
        raise ValueError(f"{name}: target must equal forecast_origin + horizon_hours")
    if result.duplicated(["plant_id", "timestamp"]).any():
        raise ValueError(f"{name}: duplicate plant/target prediction keys")
    for column in result.columns:
        if column.startswith("y_true_"):
            extra_truth = pd.to_numeric(result[column], errors="raise").to_numpy(float)
            if not np.array_equal(extra_truth, result.y_true.to_numpy(float)):
                raise ValueError(f"{name}: aligned actual values disagree in {column}")
    result["hour"] = result.timestamp.dt.hour
    return result.sort_values(["timestamp", "plant_id"], kind="stable").reset_index(drop=True)


def _metrics(actual: pd.Series, prediction: pd.Series) -> dict[str, Any]:
    y = actual.to_numpy(float)
    predicted = prediction.to_numpy(float)
    error = y - predicted
    mae = float(np.mean(np.abs(error)))
    mse = float(np.mean(np.square(error)))
    total_variance = float(np.sum(np.square(y - y.mean())))
    # Constant or singleton targets have no defined R²; never clip negative R².
    r2 = 1.0 - float(np.sum(np.square(error))) / total_variance if len(y) > 1 and total_variance > 0 else None
    if not np.isfinite([mae, mse]).all() or (r2 is not None and not np.isfinite(r2)):
        raise ValueError("Prediction magnitudes overflow finite regression metrics")
    return {"n_samples": len(y), "mae": mae, "rmse": float(np.sqrt(mse)), "r2": r2}


def _model_metrics(frame: pd.DataFrame, column: str) -> dict[str, Any]:
    plant = []
    for (plant_id, name, region), group in frame.groupby(["plant_id", "plant", "region"], sort=True):
        plant.append({"plant_id": plant_id, "plant": name, "region": region,
                      **_metrics(group.y_true, group[column])})
    region = []
    regional = frame.groupby(["region", "timestamp"], as_index=False)[["y_true", column]].sum()
    for name, group in regional.groupby("region", sort=True):
        region.append({"region": name, **_metrics(group.y_true, group[column])})
    national = frame.groupby("timestamp", as_index=False)[["y_true", column]].sum()
    return {
        "pooled": _metrics(frame.y_true, frame[column]),
        "plant": plant,
        "region": region,
        "national": _metrics(national.y_true, national[column]),
    }


def _all_metrics(frame: pd.DataFrame) -> dict[str, Any]:
    return {model: _model_metrics(frame, column) for model, column in _PREDICTIONS.items()}


def _period(frame: pd.DataFrame) -> dict[str, Any]:
    return {
        "start": frame.timestamp.min().isoformat(),
        "end": frame.timestamp.max().isoformat(),
        "forecast_origin_start": frame.forecast_origin.min().isoformat(),
        "forecast_origin_end": frame.forecast_origin.max().isoformat(),
        "rows": len(frame),
        "unique_timestamps": int(frame.timestamp.nunique()),
        "plants": int(frame.plant_id.nunique()),
    }


def _predict_gate(gate: ExplainableDynamicGate, frame: pd.DataFrame) -> pd.DataFrame:
    # Existing gate contexts key by `plant`; use the stable registry identity.
    inputs = frame.copy()
    names = frame.set_index("plant_id").plant.to_dict()
    inputs["plant"] = inputs.plant_id
    result = gate.predict(inputs).rename(columns={"selected_model": "gate_preferred_model"})
    result["plant"] = result.plant_id.map(names)
    return result


def _write_csv_atomic(frame: pd.DataFrame, path: Path, *, compression: str | None = None) -> None:
    temporary = path.with_name(path.name + ".tmp")
    try:
        frame.to_csv(temporary, index=False, compression=compression)
        replace_file_atomic(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


class BenchmarkModelSelector:
    """Fit the blend before selection, freeze the champion before reading Test errors."""

    def __init__(self, output_dir: Path, minimum_relative_improvement: float = 0.0,
                 selection_gap_hours: int = 0):
        if not np.isfinite(minimum_relative_improvement) or not 0 <= minimum_relative_improvement < 1:
            raise ValueError("minimum_relative_improvement must be finite and in [0, 1)")
        if isinstance(selection_gap_hours, bool) or int(selection_gap_hours) != selection_gap_hours or selection_gap_hours < 0:
            raise ValueError("selection_gap_hours must be a nonnegative integer")
        self.output_dir = Path(output_dir)
        self.minimum_relative_improvement = float(minimum_relative_improvement)
        self.selection_gap_hours = int(selection_gap_hours)

    def run(self, calibration: pd.DataFrame, test: pd.DataFrame, *,
            evaluation_contract: dict[str, Any], provenance: dict[str, Any]) -> dict[str, Any]:
        """Persist a selection decision and four-model errors on identical held-out rows.

        Callers tune each base model on earlier Train/Validation only and supply an
        aligned common cohort, recording discarded-row coverage in provenance.
        Calibration is split by global target time, never randomly or per plant.
        """
        horizon = evaluation_contract.get("horizon_hours")
        if isinstance(horizon, bool) or not isinstance(horizon, (int, np.integer)) or horizon <= 0:
            raise ValueError("evaluation_contract requires positive integer horizon_hours")
        # Fail early for non-JSON provenance and serialize a detached snapshot.
        contract = json.loads(json.dumps(evaluation_contract, allow_nan=False))
        lineage = json.loads(json.dumps(provenance, allow_nan=False))
        calibration = _validate_frame(calibration, "calibration", horizon)
        test = _validate_frame(test, "test", horizon)
        if str(calibration.timestamp.dt.tz) != str(test.timestamp.dt.tz):
            raise ValueError("Calibration and Test timezone conventions differ")
        if calibration.timestamp.max() >= test.forecast_origin.min():
            raise ValueError("Calibration targets must precede every Test forecast origin")
        identities = pd.concat([calibration, test]).groupby("plant_id")[["plant", "region"]].nunique()
        if identities.gt(1).any().any():
            raise ValueError("Plant identity/region mapping changes between predictions")
        times = calibration.timestamp.drop_duplicates().sort_values()
        if len(times) < 4:
            raise ValueError("Calibration needs at least four distinct target timestamps")
        selection_start = times.iloc[len(times) // 2]
        gap = max(self.selection_gap_hours, int(horizon))
        gate_fit = calibration.loc[calibration.timestamp < selection_start - pd.Timedelta(hours=gap)].copy()
        selection = calibration.loc[calibration.timestamp >= selection_start].copy()
        if gate_fit.timestamp.nunique() < 2 or selection.timestamp.nunique() < 2:
            raise ValueError("Calibration is too short after the chronological selection purge")
        if gate_fit.timestamp.max() >= selection.forecast_origin.min():
            raise ValueError("Gate-fit targets overlap selection forecast origins")
        gate_inputs = gate_fit.copy()
        gate_inputs["plant"] = gate_inputs.plant_id
        gate = ExplainableDynamicGate().fit(gate_inputs)
        selection_predictions = _predict_gate(gate, selection)
        selection_metrics = _all_metrics(selection_predictions)
        best_base = min(("xgboost", "cnn_bilstm"), key=lambda model: selection_metrics[model]["pooled"]["mae"])
        base_mae = selection_metrics[best_base]["pooled"]["mae"]
        hybrid_mae = selection_metrics["hybrid"]["pooled"]["mae"]
        improvement = (base_mae - hybrid_mae) / base_mae if base_mae > 0 else None
        threshold_mae = base_mae * (1 - self.minimum_relative_improvement)
        selected_model = "hybrid" if hybrid_mae < threshold_mae else best_base
        reason = (
            "Hybrid strictly improves held-out selection MAE beyond the required margin."
            if selected_model == "hybrid" else
            "The best base model is retained because Hybrid does not strictly improve selection MAE beyond the required margin."
        )
        # Test is consulted only after gate weights and the decision are frozen.
        test_predictions = _predict_gate(gate, test)
        test_predictions["selected_model"] = selected_model
        test_predictions["selected_pred"] = test_predictions[_PREDICTIONS[selected_model]]
        test_metrics = _all_metrics(test_predictions)
        comparison = {}
        for split, metrics in (("selection", selection_metrics), ("test", test_metrics)):
            chosen = metrics[selected_model]["pooled"]["mae"]
            persistence = metrics["persistence"]["pooled"]["mae"]
            comparison[split] = {"selected_mae": chosen, "persistence_mae": persistence,
                                 "beats_persistence": chosen < persistence}
        paths = {
            "test_predictions": self.output_dir.resolve() / "test_predictions.csv.gz",
            "gate_profiles": self.output_dir.resolve() / "gate_profiles.csv",
        }
        selection_path = self.output_dir.resolve() / "selection.json"
        if selection_path.exists() or any(path.exists() for path in paths.values()):
            raise FileExistsError("Benchmark selection artifacts already exist; use a new run directory")
        self.output_dir.mkdir(parents=True, exist_ok=True)
        _write_csv_atomic(test_predictions, paths["test_predictions"], compression="gzip")
        assert gate.profiles_ is not None
        _write_csv_atomic(gate.profiles_, paths["gate_profiles"])
        report = {
            "contract": SELECTION_CONTRACT, "schema_version": 1, "status": "completed",
            "selected_model": selected_model, "reason": reason,
            "selection_rule": {
                "metric": "pooled_mae", "best_base_model": best_base,
                "minimum_relative_improvement": self.minimum_relative_improvement,
                "hybrid_relative_improvement": improvement, "ties_favor_base": True,
                "test_used_for_selection": False,
            },
            "metric_scope": {"pooled": "aligned_plant_hour_rows",
                             "region": "aligned_cohort_hourly_generation_sums",
                             "national": "aligned_cohort_hourly_generation_sums",
                             "r2": "unclipped; null when target variance is zero"},
            "selection_metrics": selection_metrics, "test_metrics": test_metrics,
            "persistence_comparison": comparison,
            "periods": {"gate_fit": _period(gate_fit), "selection": _period(selection), "test": _period(test)},
            "purge": {"selection_gap_hours": gap, "removed_calibration_rows": len(calibration) - len(gate_fit) - len(selection)},
            "provenance": lineage, "evaluation_contract": contract,
            "files": {name: path.name for name, path in paths.items()},
            "files_sha256": {name: sha256_file(path) for name, path in paths.items()},
        }
        write_json_atomic(selection_path, report)
        return {"selection_path": str(selection_path), "selected_model": selected_model, **report}
