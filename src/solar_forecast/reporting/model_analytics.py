"""User-facing projections for model performance and solar data signals.

The dashboard deliberately ignores smoke runs and failed/incomplete artifacts.
It never promotes legacy or unmatched evaluations into a model comparison.
"""

from __future__ import annotations

from collections import Counter, defaultdict, deque
from dataclasses import dataclass
from datetime import datetime, timedelta
import heapq
import json
import math
from pathlib import Path
import sqlite3
import tempfile
from typing import Any, Iterable

import pandas as pd

from solar_forecast.collectors.normalization import read_csv_with_fallback


MODEL_LABELS = {
    "xgboost": "XGBoost",
    "cnn_bilstm": "CNN-BiLSTM",
    "hybrid": "Dynamic Hybrid",
}

CALIBRATION_CONTAMINATION = 0.05
MINIMUM_GLOBAL_CALIBRATION_SAMPLES = 5
MINIMUM_PLANT_CALIBRATION_SAMPLES = 168
TOP_EVENTS_PER_MODEL = 250
SERIES_POINTS_PER_PLANT = 168


@dataclass(frozen=True)
class _ModelRun:
    model_id: str
    manifest_path: Path
    details: dict[str, Any]


@dataclass
class _MetricAccumulator:
    n_samples: int = 0
    absolute_error: float = 0.0
    squared_error: float = 0.0
    target_sum: float = 0.0
    target_squared_sum: float = 0.0
    capacity_absolute_error: float = 0.0
    capacity_sum: float = 0.0
    capacity_samples: int = 0

    def add(self, actual: float, predicted: float, capacity_mw: float | None) -> None:
        error = actual - predicted
        self.n_samples += 1
        self.absolute_error += abs(error)
        self.squared_error += error * error
        self.target_sum += actual
        self.target_squared_sum += actual * actual
        if capacity_mw is not None and capacity_mw > 0:
            self.capacity_absolute_error += abs(error)
            self.capacity_sum += capacity_mw
            self.capacity_samples += 1

    def metrics(self) -> dict[str, int | float | None]:
        if not self.n_samples:
            return {
                "n_samples": 0,
                "mae": None,
                "rmse": None,
                "r2": None,
                "nmae_capacity": None,
                "capacity_samples": 0,
                "capacity_coverage": 0.0,
            }
        variance = self.target_squared_sum - (
            self.target_sum * self.target_sum / self.n_samples
        )
        return {
            "n_samples": self.n_samples,
            "mae": self.absolute_error / self.n_samples,
            "rmse": math.sqrt(self.squared_error / self.n_samples),
            "r2": 1.0 - self.squared_error / variance if variance > 0 else None,
            "nmae_capacity": (
                self.capacity_absolute_error / self.capacity_sum * 100.0
                if self.capacity_samples > 0 and self.capacity_sum > 0
                else None
            ),
            "capacity_samples": self.capacity_samples,
            "capacity_coverage": self.capacity_samples / self.n_samples,
        }


@dataclass(frozen=True)
class _CalibrationPolicy:
    """Frozen calibration thresholds with plant-aware, unit-safe fallbacks."""

    normalized_global: float | None
    normalized_by_plant: dict[str, float]
    absolute_global: float
    absolute_by_plant: dict[str, float]

    def evaluate(
        self,
        *,
        plant_id: str,
        absolute_error: float,
        capacity_mw: float | None,
    ) -> dict[str, Any]:
        if capacity_mw is not None and capacity_mw > 0:
            observed = absolute_error / capacity_mw
            if plant_id in self.normalized_by_plant:
                threshold = self.normalized_by_plant[plant_id]
                source = "plant_capacity_normalized"
            elif self.normalized_global is not None:
                threshold = self.normalized_global
                source = "global_capacity_normalized"
            else:
                return self._absolute_evaluation(
                    plant_id=plant_id,
                    absolute_error=absolute_error,
                    source_suffix="capacity_fallback",
                )
            return {
                "evaluable": True,
                "is_signal": observed > threshold,
                "exceedance_ratio": observed / max(threshold, 1e-12),
                "threshold_mwh": threshold * capacity_mw,
                "threshold_value": threshold,
                "threshold_unit": "capacity_ratio",
                "threshold_source": source,
                "normalized_absolute_error": observed,
            }
        return self._absolute_evaluation(
            plant_id=plant_id,
            absolute_error=absolute_error,
            source_suffix="capacity_missing",
            allow_global=False,
        )

    def _absolute_evaluation(
        self,
        *,
        plant_id: str,
        absolute_error: float,
        source_suffix: str,
        allow_global: bool = True,
    ) -> dict[str, Any]:
        if plant_id in self.absolute_by_plant:
            threshold = self.absolute_by_plant[plant_id]
            source = f"plant_absolute_{source_suffix}"
        elif allow_global:
            threshold = self.absolute_global
            source = f"global_absolute_{source_suffix}"
        else:
            return {
                "evaluable": False,
                "is_signal": False,
                "exceedance_ratio": 0.0,
                "threshold_mwh": None,
                "threshold_value": None,
                "threshold_unit": "MWh",
                "threshold_source": (
                    "deferred_capacity_missing_insufficient_plant_calibration"
                ),
                "normalized_absolute_error": None,
            }
        return {
            "evaluable": True,
            "is_signal": absolute_error > threshold,
            "exceedance_ratio": absolute_error / max(threshold, 1e-12),
            "threshold_mwh": threshold,
            "threshold_value": threshold,
            "threshold_unit": "MWh",
            "threshold_source": source,
            "normalized_absolute_error": None,
        }


class ModelAnalyticsService:
    """Build a compact public contract from approved evaluation artifacts."""

    def __init__(self, project_root: Path):
        self.project_root = Path(project_root).resolve()

    def build(self) -> dict[str, Any]:
        all_runs = self._full_runs()
        compatible_runs = self._latest_compatible_pair(all_runs)
        aligned = self._aligned_analysis(compatible_runs)
        runs = compatible_runs if aligned is not None else self._latest_runs(all_runs)
        if aligned is not None:
            models = aligned["models"]
            regions = aligned["regions"]
            plants = aligned["plants"]
            series = aligned["series"]
            prediction_signals = aligned["prediction_signals"]
            evaluation = aligned["evaluation"]
            status = "ready"
            message = "동일한 테스트 구간의 정식 평가 결과를 비교합니다."
        else:
            models = [self._model_summary(run) for run in runs]
            regions = []
            plants = []
            series = []
            prediction_signals = []
            evaluation = self._evaluation_summary(runs, set())
        if aligned is None and models:
            status = "partial"
            message = (
                "정식 결과는 있으나 동일 테스트 표본으로 정렬되지 않아 모델 간 비교는 "
                "표시하지 않습니다."
                if len(models) >= 2
                else "정식 단일 모델 결과만 있습니다. 동일 테스트 표본의 다른 모델 결과가 "
                "확보되면 모델 간 비교가 활성화됩니다."
            )
        elif aligned is None:
            status = "empty"
            message = (
                "비교 가능한 정식 평가 결과가 아직 없습니다. 동일한 테스트 구간으로 "
                "모델 평가가 완료되면 표시됩니다."
            )

        return {
            "status": status,
            "message": message,
            "evaluation": evaluation,
            "models": models,
            "regions": regions,
            "plants": plants,
            "series": series,
            "anomalies": {
                "prediction_signals": prediction_signals,
                "prediction_summary": (
                    aligned["prediction_summary"]
                    if aligned is not None
                    else self._empty_prediction_summary()
                ),
                "data_quality_signals": self._data_quality_signals(),
            },
        }

    def _aligned_analysis(self, runs: list[_ModelRun]) -> dict[str, Any] | None:
        by_model = {run.model_id: run for run in runs}
        if not {"xgboost", "cnn_bilstm"}.issubset(by_model):
            return None
        xgboost = by_model["xgboost"]
        cnn = by_model["cnn_bilstm"]
        signature = self._comparison_signature(xgboost)
        if signature is None or signature != self._comparison_signature(cnn):
            return None
        test_paths = {
            model_id: self._artifact_path(run, "test_predictions")
            for model_id, run in (("xgboost", xgboost), ("cnn_bilstm", cnn))
        }
        calibration_paths = {
            model_id: self._artifact_path(run, "calibration_predictions")
            for model_id, run in (("xgboost", xgboost), ("cnn_bilstm", cnn))
        }
        if any(
            path is None
            for path in (*test_paths.values(), *calibration_paths.values())
        ):
            return None

        try:
            with tempfile.TemporaryDirectory(
                prefix="solar-model-analytics-"
            ) as directory:
                database = Path(directory) / "predictions.sqlite3"
                connection = sqlite3.connect(database)
                try:
                    connection.execute("PRAGMA journal_mode=OFF")
                    connection.execute("PRAGMA synchronous=OFF")
                    connection.execute("PRAGMA temp_store=FILE")
                    connection.execute(
                        """
                        CREATE TABLE predictions (
                            model TEXT NOT NULL,
                            timestamp TEXT NOT NULL,
                            plant_id TEXT NOT NULL,
                            region TEXT NOT NULL,
                            plant TEXT NOT NULL,
                            y_true REAL NOT NULL,
                            y_pred REAL NOT NULL,
                            PRIMARY KEY (model, timestamp, plant_id)
                        ) WITHOUT ROWID
                        """
                    )
                    for model_id, path in test_paths.items():
                        assert path is not None
                        self._insert_predictions(connection, path, model_id)
                    connection.execute(
                        "CREATE INDEX predictions_key ON "
                        "predictions(timestamp, plant_id)"
                    )
                    capacity = self._capacity_lookup()
                    thresholds = {
                        model_id: self._calibration_threshold(
                            connection,
                            path,
                            model_id,
                            capacity,
                        )
                        for model_id, path in calibration_paths.items()
                        if path is not None
                    }
                    result = self._aggregate_aligned_rows(
                        connection,
                        xgboost,
                        thresholds,
                        capacity,
                    )
                finally:
                    connection.close()
                return result
        except (
            OSError,
            UnicodeDecodeError,
            ValueError,
            KeyError,
            sqlite3.Error,
            pd.errors.ParserError,
        ):
            # Public reporting is fail-closed: malformed or mismatched artifacts
            # remain unavailable rather than surfacing a misleading comparison.
            return None

    @staticmethod
    def _comparison_signature(run: _ModelRun) -> str | None:
        contract = run.details.get("evaluation_contract")
        if not isinstance(contract, dict):
            return None
        required = (
            "dataset_fingerprint",
            "target",
            "target_unit",
            "horizon_hours",
            "test_start",
            "test_end",
            "prediction_key",
        )
        if any(contract.get(key) in (None, "") for key in required):
            return None
        if contract.get("prediction_key") != ["timestamp", "plant_id"]:
            return None
        return json.dumps(
            {key: contract[key] for key in required},
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )

    @staticmethod
    def _artifact_path(run: _ModelRun, key: str) -> Path | None:
        value = run.details.get(key)
        if not value:
            return None
        configured = Path(str(value))
        if configured.is_file():
            return configured
        direct = run.manifest_path.parent / configured.name
        if direct.is_file():
            return direct
        matches = sorted(run.manifest_path.parent.rglob(configured.name))
        return matches[-1] if matches else None

    def _insert_predictions(
        self, connection: sqlite3.Connection, path: Path, model_id: str
    ) -> None:
        header = pd.read_csv(path, nrows=0).columns.tolist()
        prediction_column = self._prediction_column(header, model_id)
        required = {
            "timestamp",
            "plant_id",
            "region",
            "plant",
            "y_true",
            prediction_column,
        }
        missing = required - set(header)
        if missing:
            raise ValueError(
                f"Prediction artifact columns are missing: {sorted(missing)}"
            )
        selected = [*sorted(required - {prediction_column}), prediction_column]
        if "split" in header:
            selected.append("split")
        inserted = 0
        for chunk in pd.read_csv(path, usecols=selected, chunksize=100_000):
            if "split" in chunk:
                chunk = chunk.loc[chunk["split"].astype(str).str.lower().eq("test")]
            chunk = chunk.drop(columns=["split"], errors="ignore")
            if chunk.empty:
                continue
            chunk["timestamp"] = self._normalize_timestamps(chunk["timestamp"])
            chunk["plant_id"] = chunk["plant_id"].astype(str).str.strip()
            if chunk["plant_id"].isin({"", "unknown", "nan"}).any():
                raise ValueError("Prediction artifact contains an unstable plant_id")
            chunk["y_true"] = pd.to_numeric(chunk["y_true"], errors="coerce")
            chunk[prediction_column] = pd.to_numeric(
                chunk[prediction_column], errors="coerce"
            )
            numeric_finite = chunk[["y_true", prediction_column]].apply(
                lambda column: column.map(
                    lambda value: (
                        math.isfinite(float(value)) if pd.notna(value) else False
                    )
                )
            )
            if (
                chunk["timestamp"].isna().any()
                or not numeric_finite.to_numpy(dtype=bool).all()
            ):
                raise ValueError(
                    "Prediction artifact contains missing evaluation values"
                )
            records = (
                (
                    model_id,
                    str(row.timestamp),
                    str(row.plant_id),
                    str(row.region) if pd.notna(row.region) else "지역 미확인",
                    str(row.plant) if pd.notna(row.plant) else str(row.plant_id),
                    float(row.y_true),
                    float(getattr(row, prediction_column)),
                )
                for row in chunk.itertuples(index=False)
            )
            connection.executemany(
                "INSERT INTO predictions VALUES (?, ?, ?, ?, ?, ?, ?)", records
            )
            inserted += len(chunk)
            connection.commit()
        if not inserted:
            raise ValueError("Prediction artifact contains no Test rows")

    def _calibration_threshold(
        self,
        connection: sqlite3.Connection,
        path: Path,
        model_id: str,
        capacity: dict[str, float],
        contamination: float = CALIBRATION_CONTAMINATION,
    ) -> _CalibrationPolicy:
        connection.execute(
            "CREATE TABLE IF NOT EXISTS calibration_residuals ("
            "model TEXT NOT NULL, plant_id TEXT NOT NULL, "
            "basis TEXT NOT NULL, residual REAL NOT NULL)"
        )
        header = pd.read_csv(path, nrows=0).columns.tolist()
        prediction_column = self._prediction_column(header, model_id)
        required = {"plant_id", "y_true", prediction_column}
        if not required.issubset(header):
            raise ValueError(
                "Calibration artifact is missing plant identity, actual, or "
                "predicted values"
            )
        selected = ["plant_id", "y_true", prediction_column]
        if "split" in header:
            selected.append("split")
        absolute_count = 0
        for chunk in pd.read_csv(
            path, usecols=selected, chunksize=100_000
        ):
            if "split" in chunk:
                chunk = chunk.loc[
                    chunk["split"].astype(str).str.lower().eq("calibration")
                ]
            if chunk.empty:
                continue
            plant_ids = chunk["plant_id"].astype(str).str.strip()
            if plant_ids.isin({"", "unknown", "nan"}).any():
                raise ValueError("Calibration artifact contains an unstable plant_id")
            actual = pd.to_numeric(chunk["y_true"], errors="coerce")
            predicted = pd.to_numeric(chunk[prediction_column], errors="coerce")
            finite = actual.map(
                lambda value: math.isfinite(float(value)) if pd.notna(value) else False
            ) & predicted.map(
                lambda value: math.isfinite(float(value)) if pd.notna(value) else False
            )
            if not finite.to_numpy(dtype=bool).all():
                raise ValueError(
                    "Calibration artifact contains invalid evaluation values"
                )
            residuals = (actual - predicted).abs()

            def records() -> Iterable[tuple[str, str, str, float]]:
                for plant_id, residual in zip(plant_ids, residuals):
                    identifier = str(plant_id)
                    absolute = float(residual)
                    yield model_id, identifier, "absolute", absolute
                    capacity_mw = capacity.get(identifier)
                    if capacity_mw is not None and capacity_mw > 0:
                        yield (
                            model_id,
                            identifier,
                            "capacity_normalized",
                            absolute / capacity_mw,
                        )

            connection.executemany(
                "INSERT INTO calibration_residuals VALUES (?, ?, ?, ?)",
                records(),
            )
            absolute_count += len(residuals)
            connection.commit()
        if absolute_count < MINIMUM_GLOBAL_CALIBRATION_SAMPLES:
            raise ValueError("At least five calibration residuals are required")
        connection.execute(
            "CREATE INDEX IF NOT EXISTS calibration_lookup ON "
            "calibration_residuals(model, basis, plant_id, residual)"
        )
        absolute_global = self._residual_quantile(
            connection,
            model_id=model_id,
            basis="absolute",
            contamination=contamination,
            minimum_samples=MINIMUM_GLOBAL_CALIBRATION_SAMPLES,
        )
        if absolute_global is None:
            raise ValueError("Calibration threshold could not be calculated")
        return _CalibrationPolicy(
            normalized_global=self._residual_quantile(
                connection,
                model_id=model_id,
                basis="capacity_normalized",
                contamination=contamination,
                minimum_samples=MINIMUM_GLOBAL_CALIBRATION_SAMPLES,
            ),
            normalized_by_plant=self._plant_residual_quantiles(
                connection,
                model_id=model_id,
                basis="capacity_normalized",
                contamination=contamination,
            ),
            absolute_global=absolute_global,
            absolute_by_plant=self._plant_residual_quantiles(
                connection,
                model_id=model_id,
                basis="absolute",
                contamination=contamination,
            ),
        )

    @staticmethod
    def _residual_quantile(
        connection: sqlite3.Connection,
        *,
        model_id: str,
        basis: str,
        contamination: float,
        minimum_samples: int,
        plant_id: str | None = None,
    ) -> float | None:
        where = "model = ? AND basis = ?"
        values: list[Any] = [model_id, basis]
        if plant_id is not None:
            where += " AND plant_id = ?"
            values.append(plant_id)
        count_row = connection.execute(
            f"SELECT COUNT(*) FROM calibration_residuals WHERE {where}",
            values,
        ).fetchone()
        count = int(count_row[0]) if count_row else 0
        if count < minimum_samples:
            return None
        rank = min(count - 1, math.ceil((count + 1) * (1 - contamination)) - 1)
        row = connection.execute(
            f"SELECT residual FROM calibration_residuals WHERE {where} "
            "ORDER BY residual LIMIT 1 OFFSET ?",
            [*values, rank],
        ).fetchone()
        return float(row[0]) if row is not None else None

    def _plant_residual_quantiles(
        self,
        connection: sqlite3.Connection,
        *,
        model_id: str,
        basis: str,
        contamination: float,
    ) -> dict[str, float]:
        plants = connection.execute(
            "SELECT plant_id FROM calibration_residuals "
            "WHERE model = ? AND basis = ? GROUP BY plant_id "
            "HAVING COUNT(*) >= ? ORDER BY plant_id",
            (model_id, basis, MINIMUM_PLANT_CALIBRATION_SAMPLES),
        )
        result: dict[str, float] = {}
        for (plant_id,) in plants:
            threshold = self._residual_quantile(
                connection,
                model_id=model_id,
                basis=basis,
                contamination=contamination,
                minimum_samples=MINIMUM_PLANT_CALIBRATION_SAMPLES,
                plant_id=str(plant_id),
            )
            if threshold is not None:
                result[str(plant_id)] = threshold
        return result

    def _aggregate_aligned_rows(
        self,
        connection: sqlite3.Connection,
        reference_run: _ModelRun,
        thresholds: dict[str, _CalibrationPolicy],
        capacity: dict[str, float],
    ) -> dict[str, Any] | None:
        national = {
            model: _MetricAccumulator()
            for model in MODEL_LABELS
            if model != "hybrid"
        }
        regional: dict[tuple[str, str], _MetricAccumulator] = {}
        plants: dict[tuple[str, str, str, str], _MetricAccumulator] = {}
        anomaly_heaps: dict[str, list[tuple[float, int, dict[str, Any]]]] = {
            "xgboost": [],
            "cnn_bilstm": [],
        }
        recent_series: dict[str, deque[dict[str, Any]]] = defaultdict(
            lambda: deque(maxlen=SERIES_POINTS_PER_PLANT)
        )
        recent_timestamp: dict[str, datetime] = {}
        evaluated_by_model: Counter[str] = Counter()
        evaluated_by_region: Counter[str] = Counter()
        evaluated_by_plant: Counter[tuple[str, str, str]] = Counter()
        signals_by_model: Counter[str] = Counter()
        signals_by_region: Counter[str] = Counter()
        signals_by_plant: Counter[tuple[str, str, str]] = Counter()
        row_count = 0
        cursor = connection.execute(
            """
            SELECT x.timestamp, x.plant_id,
                   CASE WHEN x.region IN ('', 'unknown') THEN c.region ELSE x.region END,
                   CASE WHEN x.plant IN ('', 'unknown') THEN c.plant ELSE x.plant END,
                   x.y_true, x.y_pred, c.y_pred
              FROM predictions AS x
              JOIN predictions AS c
                ON x.timestamp = c.timestamp AND x.plant_id = c.plant_id
             WHERE x.model = 'xgboost' AND c.model = 'cnn_bilstm'
               AND ABS(x.y_true - c.y_true) <= 0.000001
             ORDER BY x.plant_id, x.timestamp
            """
        )
        while rows := cursor.fetchmany(100_000):
            for timestamp, plant_id, region, plant, actual, xgb_pred, cnn_pred in rows:
                actual = float(actual)
                plant_id = str(plant_id)
                region = str(region)
                plant = str(plant)
                predictions = {
                    "xgboost": float(xgb_pred),
                    "cnn_bilstm": float(cnn_pred),
                }
                capacity_mw = capacity.get(plant_id)
                for model_id, predicted in predictions.items():
                    national[model_id].add(actual, predicted, capacity_mw)
                    regional.setdefault(
                        (model_id, region), _MetricAccumulator()
                    ).add(actual, predicted, capacity_mw)
                    plants.setdefault(
                        (model_id, region, plant_id, plant),
                        _MetricAccumulator(),
                    ).add(actual, predicted, capacity_mw)
                    absolute_error = abs(actual - predicted)
                    evaluation = thresholds[model_id].evaluate(
                        plant_id=plant_id,
                        absolute_error=absolute_error,
                        capacity_mw=capacity_mw,
                    )
                    if not evaluation["evaluable"]:
                        continue
                    identity = (region, plant_id, plant)
                    evaluated_by_model[model_id] += 1
                    evaluated_by_region[region] += 1
                    evaluated_by_plant[identity] += 1
                    if evaluation["is_signal"]:
                        signals_by_model[model_id] += 1
                        signals_by_region[region] += 1
                        signals_by_plant[identity] += 1
                        event = {
                            "model": model_id,
                            "model_label": MODEL_LABELS[model_id],
                            "timestamp": timestamp,
                            "region": region,
                            "plant_id": plant_id,
                            "plant": plant,
                            "y_true": actual,
                            "y_pred": predicted,
                            "absolute_error": absolute_error,
                            "threshold": evaluation["threshold_mwh"],
                            "threshold_value": evaluation["threshold_value"],
                            "threshold_unit": evaluation["threshold_unit"],
                            "threshold_source": evaluation["threshold_source"],
                            "normalized_absolute_error": evaluation[
                                "normalized_absolute_error"
                            ],
                            "exceedance_ratio": evaluation["exceedance_ratio"],
                        }
                        heap = anomaly_heaps[model_id]
                        candidate = (
                            float(evaluation["exceedance_ratio"]),
                            row_count,
                            event,
                        )
                        if len(heap) < TOP_EVENTS_PER_MODEL:
                            heapq.heappush(heap, candidate)
                        elif candidate[0] > heap[0][0]:
                            heapq.heapreplace(heap, candidate)
                sample = {
                    "timestamp": timestamp,
                    "plant_id": plant_id,
                    "region": region,
                    "plant": plant,
                    "y_true": actual,
                    "predictions": predictions,
                }
                parsed_timestamp = datetime.fromisoformat(timestamp)
                previous_timestamp = recent_timestamp.get(plant_id)
                if (
                    previous_timestamp is not None
                    and parsed_timestamp - previous_timestamp != timedelta(hours=1)
                ):
                    recent_series[plant_id].clear()
                recent_series[plant_id].append(sample)
                recent_timestamp[plant_id] = parsed_timestamp
                row_count += 1
        if not row_count:
            return None

        model_rows = [
            {
                "id": model_id,
                "label": MODEL_LABELS[model_id],
                "comparable": True,
                "metrics": national[model_id].metrics(),
            }
            for model_id in ("xgboost", "cnn_bilstm")
        ]
        region_rows = [
            {"model": model, "region": region, "metrics": accumulator.metrics()}
            for (model, region), accumulator in regional.items()
        ]
        plant_rows = [
            {
                "model": model,
                "region": region,
                "plant_id": plant_id,
                "plant": plant,
                "metrics": accumulator.metrics(),
            }
            for (model, region, plant_id, plant), accumulator in plants.items()
        ]
        events = [item[2] for heap in anomaly_heaps.values() for item in heap]
        events.sort(key=lambda item: item["exceedance_ratio"], reverse=True)
        series = [item for values in recent_series.values() for item in values]
        prediction_summary = self._prediction_summary(
            evaluated_by_model=evaluated_by_model,
            evaluated_by_region=evaluated_by_region,
            evaluated_by_plant=evaluated_by_plant,
            signals_by_model=signals_by_model,
            signals_by_region=signals_by_region,
            signals_by_plant=signals_by_plant,
            returned_top_events=len(events),
        )
        contract = reference_run.details["evaluation_contract"]
        return {
            "models": model_rows,
            "regions": sorted(
                region_rows, key=lambda item: (item["region"], item["model"])
            ),
            "plants": sorted(
                plant_rows,
                key=lambda item: (item["region"], item["plant"], item["model"]),
            ),
            "series": sorted(
                series, key=lambda item: (item["timestamp"], item["plant_id"])
            ),
            "prediction_signals": events,
            "prediction_summary": prediction_summary,
            "evaluation": {
                "scope": "test",
                "from": contract["test_start"],
                "to": contract["test_end"],
                "horizon_hours": int(contract["horizon_hours"]),
                "common_samples": row_count,
            },
        }

    @staticmethod
    def _prediction_summary(
        *,
        evaluated_by_model: Counter[str],
        evaluated_by_region: Counter[str],
        evaluated_by_plant: Counter[tuple[str, str, str]],
        signals_by_model: Counter[str],
        signals_by_region: Counter[str],
        signals_by_plant: Counter[tuple[str, str, str]],
        returned_top_events: int,
    ) -> dict[str, Any]:
        total_evaluated = sum(evaluated_by_model.values())
        total_signals = sum(signals_by_model.values())

        def rate(signals: int, evaluated: int) -> float:
            return signals / evaluated if evaluated else 0.0

        return {
            "total": total_signals,
            "evaluated_predictions": total_evaluated,
            "rate": rate(total_signals, total_evaluated),
            "top_event_limit_per_model": TOP_EVENTS_PER_MODEL,
            "returned_top_events": returned_top_events,
            "by_model": [
                {
                    "model": model_id,
                    "model_label": MODEL_LABELS[model_id],
                    "signals": signals_by_model[model_id],
                    "evaluated_predictions": evaluated_by_model[model_id],
                    "rate": rate(
                        signals_by_model[model_id], evaluated_by_model[model_id]
                    ),
                }
                for model_id in ("xgboost", "cnn_bilstm")
            ],
            "by_region": sorted(
                (
                    {
                        "region": region,
                        "signals": signals_by_region[region],
                        "evaluated_predictions": evaluated,
                        "rate": rate(signals_by_region[region], evaluated),
                    }
                    for region, evaluated in evaluated_by_region.items()
                ),
                key=lambda item: (-item["signals"], item["region"]),
            ),
            "by_plant": sorted(
                (
                    {
                        "region": region,
                        "plant_id": plant_id,
                        "plant": plant,
                        "signals": signals_by_plant[(region, plant_id, plant)],
                        "evaluated_predictions": evaluated,
                        "rate": rate(
                            signals_by_plant[(region, plant_id, plant)], evaluated
                        ),
                    }
                    for (region, plant_id, plant), evaluated in (
                        evaluated_by_plant.items()
                    )
                ),
                key=lambda item: (
                    -item["signals"],
                    item["region"],
                    item["plant"],
                    item["plant_id"],
                ),
            ),
        }

    @staticmethod
    def _empty_prediction_summary() -> dict[str, Any]:
        return {
            "total": 0,
            "evaluated_predictions": 0,
            "rate": 0.0,
            "top_event_limit_per_model": TOP_EVENTS_PER_MODEL,
            "returned_top_events": 0,
            "by_model": [],
            "by_region": [],
            "by_plant": [],
        }

    def _capacity_lookup(self) -> dict[str, float]:
        path = self.project_root / "file/standardized/plant_registry.csv"
        if not path.exists():
            return {}
        frame = read_csv_with_fallback(path)
        if not {"plant_id", "capacity_mw"}.issubset(frame.columns):
            return {}
        frame = frame[["plant_id", "capacity_mw"]]
        identifiers = frame["plant_id"].astype(str).str.strip()
        numeric = pd.to_numeric(frame["capacity_mw"], errors="coerce")
        return {
            str(plant_id): float(value)
            for plant_id, value in zip(identifiers, numeric)
            if plant_id not in {"", "unknown", "nan"}
            and pd.notna(value)
            and math.isfinite(float(value))
            and float(value) > 0
        }

    @staticmethod
    def _prediction_column(columns: Iterable[str], model_id: str) -> str:
        candidates = (
            "y_pred",
            "xgb_pred" if model_id == "xgboost" else "cnn_pred",
        )
        return next(
            (column for column in candidates if column in columns), candidates[-1]
        )

    @staticmethod
    def _normalize_timestamps(values: pd.Series) -> pd.Series:
        parsed = pd.to_datetime(values, errors="coerce")
        return parsed.dt.strftime("%Y-%m-%dT%H:%M:%S")

    def _full_runs(self) -> list[_ModelRun]:
        runs: list[_ModelRun] = []
        for model_id in ("xgboost", "cnn_bilstm"):
            root = self.project_root / "artifacts/models" / model_id
            if root.exists():
                for path in root.glob("*/manifest.json"):
                    manifest = self._read_json(path)
                    details = manifest.get("details")
                    if (
                        manifest.get("status") != "completed"
                        or manifest.get("model") != model_id
                        or not isinstance(details, dict)
                        or not self._is_full_run(details)
                    ):
                        continue
                    runs.append(_ModelRun(model_id, path, details))
        return runs

    @staticmethod
    def _latest_runs(runs: list[_ModelRun]) -> list[_ModelRun]:
        latest: list[_ModelRun] = []
        for model_id in ("xgboost", "cnn_bilstm"):
            candidates = [run for run in runs if run.model_id == model_id]
            if candidates:
                latest.append(
                    max(candidates, key=lambda run: run.manifest_path.parent.name)
                )
        return latest

    def _latest_compatible_pair(
        self, runs: list[_ModelRun]
    ) -> list[_ModelRun]:
        grouped: dict[str, dict[str, list[_ModelRun]]] = defaultdict(
            lambda: defaultdict(list)
        )
        for run in runs:
            signature = self._comparison_signature(run)
            if signature is not None:
                grouped[signature][run.model_id].append(run)

        pairs: list[tuple[tuple[str, str], _ModelRun, _ModelRun]] = []
        for models in grouped.values():
            if not {"xgboost", "cnn_bilstm"}.issubset(models):
                continue
            xgboost = max(
                models["xgboost"], key=lambda run: run.manifest_path.parent.name
            )
            cnn = max(
                models["cnn_bilstm"], key=lambda run: run.manifest_path.parent.name
            )
            names = sorted(
                (xgboost.manifest_path.parent.name, cnn.manifest_path.parent.name)
            )
            # A comparison only becomes current when both sides exist. Prefer
            # the pair with the newest older member, then the newest newer
            # member, rather than letting one fresh but stale-paired run win.
            pairs.append(((names[0], names[1]), xgboost, cnn))
        if not pairs:
            return []
        _, xgboost, cnn = max(pairs, key=lambda item: item[0])
        return [xgboost, cnn]

    @staticmethod
    def _is_full_run(details: dict[str, Any]) -> bool:
        context = details.get("run_context")
        if not isinstance(context, dict):
            # Older manifests overwrote their smoke flag on completion and are
            # therefore not trustworthy enough for a public comparison.
            return False
        return (
            context.get("execution_mode") == "full"
            and str(context.get("energy_source", "")).strip().lower() == "solar"
            and context.get("target") == "generation_mwh"
            and context.get("target_unit") == "MWh"
        )

    def _model_summary(self, run: _ModelRun) -> dict[str, Any]:
        metrics = run.details.get("metrics")
        metrics = metrics if isinstance(metrics, dict) else {}
        split = run.details.get("temporal_split")
        split = split if isinstance(split, dict) else {}
        counts = split.get("counts")
        counts = counts if isinstance(counts, dict) else {}
        n_samples = run.details.get("n_test", counts.get("test"))
        return {
            "id": run.model_id,
            "label": MODEL_LABELS[run.model_id],
            "comparable": False,
            "metrics": {
                "n_samples": self._integer(n_samples),
                "mae": self._number(metrics.get("mae")),
                "rmse": self._number(metrics.get("rmse")),
                "r2": self._number(metrics.get("r2")),
                # Capacity-normalized MAE must be calculated from aligned rows
                # and capacity metadata; never infer it from a national mean.
                "nmae_capacity": None,
                "capacity_samples": None,
                "capacity_coverage": None,
            },
        }

    def _evaluation_summary(
        self, runs: list[_ModelRun], comparable: set[str]
    ) -> dict[str, Any]:
        selected = next((run for run in runs if run.model_id in comparable), None)
        contract = selected.details.get("evaluation_contract", {}) if selected else {}
        if not isinstance(contract, dict):
            contract = {}
        samples = [
            self._model_summary(run)["metrics"]["n_samples"]
            for run in runs
            if run.model_id in comparable
        ]
        samples = [value for value in samples if value is not None]
        return {
            "scope": "test",
            "from": contract.get("test_start"),
            "to": contract.get("test_end"),
            "horizon_hours": self._integer(contract.get("horizon_hours")),
            # Until a dedicated aligned evaluator writes the exact intersection,
            # the safe lower bound is used only for display metadata.
            "common_samples": min(samples) if samples else 0,
        }

    def _data_quality_signals(self) -> list[dict[str, Any]]:
        path = self.project_root / "file/standardized/plant_quality_report.csv"
        if not path.exists():
            return []
        frame = read_csv_with_fallback(path)
        required = {"plant", "region", "energy_source", "sensor_risk"}
        if not required.issubset(frame.columns):
            return []
        solar = frame.loc[
            frame["energy_source"].astype(str).str.lower().eq("solar")
            & frame["sensor_risk"].astype(str).str.lower().isin(
                {"review", "high", "pipeline_artifact"}
            )
        ].copy()
        signals = [self._quality_signal(row) for _, row in solar.iterrows()]
        return sorted(signals, key=lambda item: (item["region"], item["plant"]))

    def _quality_signal(self, row: pd.Series) -> dict[str, Any]:
        values = {
            key: self._number(row.get(key))
            for key in (
                "hourly_coverage",
                "missing_weather_rate",
                "daylight_zero_rate",
                "capacity_exceeded_rate",
                "positive_flatline_rate",
                "temporal_profile_consistency",
                "peer_pattern_correlation",
            )
        }
        labels: list[str] = []
        if self._at_least(values["missing_weather_rate"], 0.30):
            labels.append("기상 관측 공백")
        if self._at_least(values["daylight_zero_rate"], 0.30):
            labels.append("주간 0발전 패턴")
        if self._at_least(values["capacity_exceeded_rate"], 0.02):
            labels.append("설비용량 초과 패턴")
        if self._at_least(values["positive_flatline_rate"], 0.01):
            labels.append("동일값 지속 패턴")
        if self._below(values["temporal_profile_consistency"], 0.80):
            labels.append("시간대 패턴 차이")
        if self._below(values["peer_pattern_correlation"], 0.20):
            labels.append("동일지역 패턴 차이")
        if not labels:
            labels.append("발전량 패턴 검토")
        return {
            "plant_id": self._text(row.get("plant_id")),
            "region": self._text(row.get("region")) or "지역 미확인",
            "plant": self._text(row.get("plant")) or "발전소 미확인",
            "severity": "review",
            "signal_types": labels,
            "summary": " · ".join(labels),
            "observations": self._integer(row.get("rows")),
            "period": {
                "from": self._text(row.get("start")),
                "to": self._text(row.get("end")),
            },
            "metrics": values,
        }

    @staticmethod
    def _read_json(path: Path) -> dict[str, Any]:
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError):
            return {}
        return payload if isinstance(payload, dict) else {}

    @staticmethod
    def _text(value: object) -> str | None:
        if value is None or pd.isna(value):
            return None
        text = str(value).strip()
        return text or None

    @staticmethod
    def _number(value: object) -> float | None:
        try:
            number = float(value)
        except (TypeError, ValueError):
            return None
        return number if math.isfinite(number) else None

    @staticmethod
    def _integer(value: object) -> int | None:
        number = ModelAnalyticsService._number(value)
        return int(number) if number is not None else None

    @staticmethod
    def _at_least(value: float | None, threshold: float) -> bool:
        return value is not None and value >= threshold

    @staticmethod
    def _below(value: float | None, threshold: float) -> bool:
        return value is not None and value < threshold


__all__ = ["ModelAnalyticsService"]
