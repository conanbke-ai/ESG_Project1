from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Sequence

from xgboost import XGBRegressor
from xgboost.callback import TrainingCallback

from .checkpointing import TrainingCheckpointStore


class ResumableEarlyStoppingCallback(TrainingCallback):
    """Early stopping whose best score and wait counter live in the Booster."""

    def __init__(self, rounds: int, metric_name: str):
        if rounds < 1:
            raise ValueError("XGBoost early_stopping_rounds must be positive")
        self.rounds = rounds
        self.metric_name = metric_name
        self.best_score = float("inf")
        self.wait = 0

    def before_training(self, model: Any) -> Any:
        best_score = model.attr("solar_early_stopping_best_score")
        wait = model.attr("solar_early_stopping_wait")
        if best_score is not None:
            self.best_score = float(best_score)
        if wait is not None:
            self.wait = int(wait)
        return model

    def after_iteration(self, model: Any, epoch: int, evals_log: dict) -> bool:
        validation = evals_log.get("validation_0", {})
        values = validation.get(self.metric_name)
        if not values:
            raise ValueError(
                f"Validation metric {self.metric_name!r} is required for early stopping"
            )
        score = float(values[-1])
        iteration = int(model.num_boosted_rounds()) - 1
        if score + 1e-12 < self.best_score:
            self.best_score = score
            self.wait = 0
            best_iteration = iteration
        else:
            self.wait += 1
            stored_iteration = model.attr("solar_early_stopping_best_iteration")
            best_iteration = (
                int(stored_iteration) if stored_iteration is not None else iteration
            )
        model.set_attr(
            best_score=str(self.best_score),
            best_iteration=str(best_iteration),
            solar_early_stopping_best_score=str(self.best_score),
            solar_early_stopping_best_iteration=str(best_iteration),
            solar_early_stopping_wait=str(self.wait),
        )
        return self.wait >= self.rounds


class XGBoostCheckpointCallback(TrainingCallback):
    """Persist a Booster atomically every configured number of rounds."""

    def __init__(
        self,
        store: TrainingCheckpointStore,
        *,
        stage: str,
        signature: str,
        initial_rounds: int = 0,
    ):
        self.store = store
        self.stage = stage
        self.signature = signature
        self.initial_rounds = initial_rounds

    def after_iteration(self, model: Any, epoch: int, evals_log: dict) -> bool:
        completed_rounds = self.initial_rounds + epoch + 1
        if completed_rounds % self.store.xgboost_every_rounds == 0:
            self.store.save_xgboost(
                model,
                self.stage,
                signature=self.signature,
                completed_rounds=completed_rounds,
                completed=False,
            )
        return False


@dataclass(frozen=True)
class ResumableXGBoostFit:
    model: XGBRegressor
    resumed: bool
    initial_rounds: int
    completed_rounds: int
    checkpoint_stage: str


def fit_xgboost_resumable(
    params: dict[str, Any],
    train_x: Any,
    train_y: Any,
    validation_x: Any,
    validation_y: Any,
    *,
    store: TrainingCheckpointStore | None,
    stage: str,
    signature: str,
    verbose: bool = False,
    callbacks: Sequence[TrainingCallback] = (),
) -> ResumableXGBoostFit:
    """Continue boosting from a compatible periodic checkpoint when present."""

    configured_rounds = int(params.get("n_estimators", 100))
    if configured_rounds < 1:
        raise ValueError("XGBoost n_estimators must be positive")
    checkpoint_path = None
    checkpoint_completed = False
    initial_rounds = 0
    resumed = False
    if store is not None:
        loaded = store.load_xgboost(stage, signature=signature)
        if loaded is not None:
            checkpoint_path, metadata = loaded
            saved_model = XGBRegressor()
            saved_model.load_model(checkpoint_path)
            initial_rounds = int(saved_model.get_booster().num_boosted_rounds())
            checkpoint_completed = bool(metadata.get("completed", False))
            resumed = initial_rounds > 0

    if checkpoint_path is not None and (
        checkpoint_completed or initial_rounds >= configured_rounds
    ):
        model = XGBRegressor()
        model.load_model(checkpoint_path)
        return ResumableXGBoostFit(
            model=model,
            resumed=resumed,
            initial_rounds=initial_rounds,
            completed_rounds=int(model.get_booster().num_boosted_rounds()),
            checkpoint_stage=stage,
        )

    fit_params = dict(params)
    fit_params["n_estimators"] = max(1, configured_rounds - initial_rounds)
    fit_callbacks = list(fit_params.pop("callbacks", ()))
    fit_callbacks.extend(callbacks)
    early_stopping_rounds = fit_params.pop("early_stopping_rounds", None)
    if early_stopping_rounds is not None:
        eval_metric = fit_params.get("eval_metric", "mae")
        if not isinstance(eval_metric, str):
            raise ValueError("Resumable XGBoost early stopping requires one metric name")
        fit_callbacks.append(
            ResumableEarlyStoppingCallback(
                int(early_stopping_rounds),
                metric_name=eval_metric,
            )
        )
    if store is not None and store.enabled:
        fit_callbacks.append(
            XGBoostCheckpointCallback(
                store,
                stage=stage,
                signature=signature,
                initial_rounds=initial_rounds,
            )
        )
    if fit_callbacks:
        fit_params["callbacks"] = fit_callbacks
    model = XGBRegressor(**fit_params)
    model.fit(
        train_x,
        train_y,
        eval_set=[(validation_x, validation_y)],
        verbose=verbose,
        xgb_model=str(checkpoint_path) if checkpoint_path is not None else None,
    )
    completed_rounds = int(model.get_booster().num_boosted_rounds())
    if store is not None:
        store.save_xgboost(
            model.get_booster(),
            stage,
            signature=signature,
            completed_rounds=completed_rounds,
            completed=True,
        )
    return ResumableXGBoostFit(
        model=model,
        resumed=resumed,
        initial_rounds=initial_rounds,
        completed_rounds=completed_rounds,
        checkpoint_stage=stage,
    )
