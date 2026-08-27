from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Mapping

import optuna
from optuna.study import Study
from optuna.trial import FrozenTrial, Trial, TrialState

from solar_forecast.artifacts.manifest import write_json_atomic
from solar_forecast.settings import PROJECT_ROOT


@dataclass(frozen=True)
class OptimizationSettings:
    """Shared, bounded Optuna contract for independently trained models."""

    enabled: bool
    study_name: str
    storage_path: Path
    max_trials: int
    timeout_seconds: int | None
    seed: int
    startup_trials: int
    pruner_startup_trials: int
    pruner_warmup_steps: int
    objective_metric: str = "validation_mae"

    @classmethod
    def from_values(
        cls,
        values: Mapping[str, object],
        *,
        model: str,
    ) -> "OptimizationSettings":
        raw = values.get("optimizer", {})
        if not isinstance(raw, Mapping):
            raise ValueError("optimizer configuration must be an object")
        feature_contract = str(values.get("feature_contract", "default"))
        timeout_value = raw.get("timeout_seconds", 3600)
        timeout_seconds = None if timeout_value is None else int(timeout_value)
        settings = cls(
            enabled=bool(raw.get("enabled", values.get("use_optuna", False))),
            study_name=str(
                raw.get("study_name", f"{model}_{feature_contract}_solar_v1")
            ),
            storage_path=Path(
                str(raw.get("storage_path", "artifacts/optimization/solar_models.db"))
            ),
            max_trials=int(raw.get("max_trials", values.get("n_trials", 10))),
            timeout_seconds=timeout_seconds,
            seed=int(raw.get("seed", values.get("seed", 42))),
            startup_trials=int(raw.get("startup_trials", 5)),
            pruner_startup_trials=int(raw.get("pruner_startup_trials", 5)),
            pruner_warmup_steps=int(raw.get("pruner_warmup_steps", 5)),
            objective_metric=str(raw.get("objective_metric", "validation_mae")),
        )
        settings.validate()
        return settings

    def validate(self) -> None:
        if not self.study_name.strip():
            raise ValueError("optimizer study_name cannot be blank")
        if self.max_trials < 1:
            raise ValueError("optimizer max_trials must be positive")
        if self.timeout_seconds is not None and self.timeout_seconds < 1:
            raise ValueError("optimizer timeout_seconds must be positive or null")
        if min(
            self.startup_trials,
            self.pruner_startup_trials,
            self.pruner_warmup_steps,
        ) < 0:
            raise ValueError("optimizer startup and warmup values cannot be negative")
        if self.objective_metric != "validation_mae":
            raise ValueError(
                "optimizer objective_metric must be validation_mae; Test is reserved"
            )


@dataclass(frozen=True)
class OptimizationRun:
    study: Study
    summary_path: Path
    trials_path: Path
    existing_trials: int
    executed_trials: int

    @property
    def best_params(self) -> dict[str, object]:
        return dict(self.study.best_params)


class OptunaStudyService:
    """Run or resume one versioned study and persist auditable artifacts."""

    def __init__(
        self,
        settings: OptimizationSettings,
        *,
        project_root: Path = PROJECT_ROOT,
    ):
        self.settings = settings
        self.project_root = Path(project_root)

    def run(
        self,
        objective: Callable[[Trial], float],
        artifact_dir: Path,
        *,
        baseline_params: Mapping[str, object] | None = None,
    ) -> OptimizationRun:
        storage_path = self.settings.storage_path
        if not storage_path.is_absolute():
            storage_path = self.project_root / storage_path
        storage_path.parent.mkdir(parents=True, exist_ok=True)
        storage = optuna.storages.RDBStorage(
            url=f"sqlite:///{storage_path.resolve().as_posix()}"
        )
        sampler = optuna.samplers.TPESampler(
            seed=self.settings.seed,
            n_startup_trials=self.settings.startup_trials,
        )
        pruner = optuna.pruners.MedianPruner(
            n_startup_trials=self.settings.pruner_startup_trials,
            n_warmup_steps=self.settings.pruner_warmup_steps,
        )
        study = optuna.create_study(
            study_name=self.settings.study_name,
            storage=storage,
            load_if_exists=True,
            direction="minimize",
            sampler=sampler,
            pruner=pruner,
        )
        baseline_enqueued = False
        if not study.trials and baseline_params:
            study.enqueue_trial(
                dict(baseline_params),
                user_attrs={"parameter_source": "previous_optuna_best_baseline"},
            )
            baseline_enqueued = True
        existing = self._finished_trials(study.trials)
        remaining = max(0, self.settings.max_trials - existing)
        if remaining:
            study.optimize(
                objective,
                n_trials=remaining,
                timeout=self.settings.timeout_seconds,
                gc_after_trial=True,
            )
        executed = self._finished_trials(study.trials) - existing
        completed = [
            trial for trial in study.trials if trial.state == TrialState.COMPLETE
        ]
        if not completed:
            raise RuntimeError(
                f"Optuna study {self.settings.study_name!r} has no completed trial"
            )

        artifact_dir = Path(artifact_dir)
        artifact_dir.mkdir(parents=True, exist_ok=True)
        trials_path = artifact_dir / "optimization_trials.csv"
        temporary = trials_path.with_name(trials_path.name + ".tmp")
        study.trials_dataframe().to_csv(temporary, index=False, encoding="utf-8-sig")
        temporary.replace(trials_path)
        summary_path = artifact_dir / "optimization_summary.json"
        write_json_atomic(
            summary_path,
            {
                "study_name": self.settings.study_name,
                "storage_path": str(storage_path),
                "direction": "minimize",
                "selection_data": "validation_only",
                "objective_metric": self.settings.objective_metric,
                "max_total_trials": self.settings.max_trials,
                "existing_finished_trials": existing,
                "baseline_enqueued": baseline_enqueued,
                "baseline_params": dict(baseline_params or {}),
                "executed_trials": executed,
                "finished_trials": self._finished_trials(study.trials),
                "completed_trials": len(completed),
                "pruned_trials": sum(
                    trial.state == TrialState.PRUNED for trial in study.trials
                ),
                "best_trial_number": study.best_trial.number,
                "best_validation_mae": float(study.best_value),
                "best_params": dict(study.best_params),
                "test_usage": "none",
            },
        )
        return OptimizationRun(
            study=study,
            summary_path=summary_path,
            trials_path=trials_path,
            existing_trials=existing,
            executed_trials=executed,
        )

    @staticmethod
    def _finished_trials(trials: list[FrozenTrial]) -> int:
        return sum(
            trial.state in {TrialState.COMPLETE, TrialState.PRUNED, TrialState.FAIL}
            for trial in trials
        )
