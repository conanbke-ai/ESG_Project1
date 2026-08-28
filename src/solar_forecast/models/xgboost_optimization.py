from __future__ import annotations

from dataclasses import dataclass
import gc
from pathlib import Path
from typing import Mapping, Sequence

import numpy as np
import optuna
import pandas as pd
from sklearn.metrics import mean_absolute_error
from xgboost.callback import TrainingCallback

from .checkpointing import TrainingCheckpointStore, stable_signature
from .optimization import OptimizationRun, OptimizationSettings, OptunaStudyService
from .xgboost_checkpoint import fit_xgboost_resumable


class OptunaPruningCallback(TrainingCallback):
    """Report XGBoost validation MAE each boosting round to Optuna."""

    def __init__(self, trial: optuna.Trial):
        self.trial = trial

    def after_iteration(self, model, epoch: int, evals_log: dict) -> bool:
        validation = evals_log.get("validation_0", {})
        values = validation.get("mae")
        if not values:
            return False
        completed_round = max(epoch, int(model.num_boosted_rounds()) - 1)
        self.trial.report(float(values[-1]), step=completed_round)
        if self.trial.should_prune():
            raise optuna.TrialPruned()
        return False


@dataclass(frozen=True)
class XGBoostOptimizationResult:
    best_params: dict[str, object]
    best_iteration: int | None
    tuning_train_rows: int
    tuning_validation_rows: int
    run: OptimizationRun

    def to_dict(self) -> dict[str, object]:
        return {
            "enabled": True,
            "selection_data": "validation_only",
            "objective_metric": "validation_mae",
            "best_validation_mae": float(self.run.study.best_value),
            "best_params": self.best_params,
            "best_iteration": self.best_iteration,
            "tuning_train_rows": self.tuning_train_rows,
            "tuning_validation_rows": self.tuning_validation_rows,
            "existing_trials": self.run.existing_trials,
            "executed_trials": self.run.executed_trials,
            "summary_path": str(self.run.summary_path),
            "trials_path": str(self.run.trials_path),
            "test_usage": "none",
        }


class XGBoostHyperparameterOptimizer:
    """Bounded multi-fidelity search followed by full-data model fitting."""

    def __init__(
        self,
        settings: OptimizationSettings,
        values: Mapping[str, object],
        checkpoint_store: TrainingCheckpointStore | None = None,
    ):
        self.settings = settings
        raw = values.get("optimizer", {})
        if not isinstance(raw, Mapping):
            raise ValueError("optimizer configuration must be an object")
        self.values = values
        self.raw = raw
        self.checkpoint_store = checkpoint_store

    def optimize(
        self,
        train_frame: pd.DataFrame,
        validation_frame: pd.DataFrame,
        *,
        feature_columns: Sequence[str],
        target_column: str,
        artifact_dir: Path,
    ) -> XGBoostOptimizationResult:
        train = self._bounded_rows(
            train_frame,
            int(self.raw.get("tuning_train_max_rows", 750_000)),
        )
        validation = self._bounded_rows(
            validation_frame,
            int(self.raw.get("tuning_validation_max_rows", 250_000)),
        )
        max_estimators = int(self.raw.get("trial_max_estimators", 1_500))
        early_stopping_rounds = int(
            self.raw.get(
                "early_stopping_rounds",
                self.values.get("early_stopping_rounds", 30),
            )
        )
        if min(max_estimators, early_stopping_rounds) < 1:
            raise ValueError("XGBoost optimizer estimator and early-stop values must be positive")

        def objective(trial: optuna.Trial) -> float:
            params = {
                "n_estimators": max_estimators,
                "max_depth": trial.suggest_int("max_depth", 3, 12),
                "learning_rate": trial.suggest_float(
                    "learning_rate", 0.01, 0.20, log=True
                ),
                "min_child_weight": trial.suggest_float(
                    "min_child_weight", 1.0, 32.0, log=True
                ),
                "subsample": trial.suggest_float("subsample", 0.60, 1.0),
                "colsample_bytree": trial.suggest_float(
                    "colsample_bytree", 0.60, 1.0
                ),
                "reg_alpha": trial.suggest_float(
                    "reg_alpha", 1e-8, 10.0, log=True
                ),
                "reg_lambda": trial.suggest_float(
                    "reg_lambda", 1e-3, 100.0, log=True
                ),
                "gamma": trial.suggest_float("gamma", 1e-8, 5.0, log=True),
                "max_bin": trial.suggest_categorical("max_bin", [128, 256, 512]),
                "tree_method": "hist",
                "eval_metric": "mae",
                "early_stopping_rounds": early_stopping_rounds,
                "random_state": self.settings.seed,
                "n_jobs": int(self.values.get("n_jobs", 4)),
            }
            checkpoint_stage = "optuna_trial_" + stable_signature(
                {"study_name": self.settings.study_name, "params": trial.params}
            )[:20]
            checkpoint_signature = stable_signature(
                {
                    "feature_columns": list(feature_columns),
                    "target_column": target_column,
                    "params": {
                        key: value for key, value in params.items() if key != "n_jobs"
                    },
                }
            )
            if self.checkpoint_store is not None:
                trial.set_user_attr("checkpoint_stage", checkpoint_stage)
            model = None
            try:
                fit = fit_xgboost_resumable(
                    params,
                    train[list(feature_columns)],
                    train[target_column],
                    validation[list(feature_columns)],
                    validation[target_column],
                    store=self.checkpoint_store,
                    stage=checkpoint_stage,
                    signature=checkpoint_signature,
                    verbose=False,
                    callbacks=[OptunaPruningCallback(trial)],
                )
                model = fit.model
                prediction = model.predict(validation[list(feature_columns)])
                score = float(
                    mean_absolute_error(validation[target_column], prediction)
                )
                try:
                    best_iteration = int(model.best_iteration)
                except (AttributeError, ValueError):
                    best_iteration = fit.completed_rounds - 1
                trial.set_user_attr("best_iteration", best_iteration)
                trial.set_user_attr("checkpoint_resumed", fit.resumed)
                trial.set_user_attr("tuning_train_rows", len(train))
                trial.set_user_attr("tuning_validation_rows", len(validation))
                return score
            finally:
                if model is not None:
                    del model
                gc.collect()

        def cleanup_checkpoint(_study: optuna.Study, frozen_trial) -> None:
            if self.checkpoint_store is None:
                return
            stage = frozen_trial.user_attrs.get("checkpoint_stage")
            if stage and frozen_trial.state in {
                optuna.trial.TrialState.COMPLETE,
                optuna.trial.TrialState.PRUNED,
            }:
                self.checkpoint_store.remove(str(stage), kind="xgboost")

        run = OptunaStudyService(self.settings).run(
            objective,
            artifact_dir,
            callbacks=[cleanup_checkpoint],
        )
        best_iteration = run.study.best_trial.user_attrs.get("best_iteration")
        return XGBoostOptimizationResult(
            best_params=dict(run.study.best_params),
            best_iteration=int(best_iteration) if best_iteration is not None else None,
            tuning_train_rows=len(train),
            tuning_validation_rows=len(validation),
            run=run,
        )

    @staticmethod
    def _bounded_rows(frame: pd.DataFrame, maximum: int) -> pd.DataFrame:
        if maximum < 1:
            raise ValueError("optimizer tuning row limits must be positive")
        if len(frame) <= maximum:
            return frame
        positions = np.linspace(0, len(frame) - 1, num=maximum, dtype=np.int64)
        return frame.iloc[positions]
