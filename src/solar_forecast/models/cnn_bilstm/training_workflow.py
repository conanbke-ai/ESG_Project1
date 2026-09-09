"""CNN-BiLSTM training workflow and its artifact contract."""
from __future__ import annotations
import json
import random
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Sequence
import pandas as pd
import numpy as np
import torch
from solar_forecast.models.shared.optuna_study import OptimizationSettings
from solar_forecast.infrastructure.artifact_store import replace_file_atomic
from solar_forecast.models.shared.checkpoint_store import (
    CHECKPOINT_CONTRACT,
    TrainingCheckpointStore,
    capture_rng_state,
    dataframe_signature,
    restore_rng_state,
    stable_signature,
)
from solar_forecast.models.cnn_bilstm.sequence_data import (
    SequenceConfig,
    prepare_dataset_splits,
    prepare_datasets,
)
from solar_forecast.models.cnn_bilstm.network import (
    CnnBiLstmNetworkConfig,
    build_cnn_bilstm_network,
)
from solar_forecast.models.cnn_bilstm.optimization import (
    evaluate_cnn_bilstm_loader,
    train_cnn_bilstm_epoch,
    save_study_results,
    train_with_best_trial,
)
from solar_forecast.models.cnn_bilstm.adaptive_training import BanditConfig, run_adaptive_training
from solar_forecast.infrastructure.artifact_store import create_run_directory


def _write_prediction_artifact(
    model: torch.nn.Module,
    loader,
    path: Path,
    *,
    split: str,
    device: torch.device,
) -> Path:
    """Stream row-aligned CNN predictions without materializing a full table."""

    dataset = loader.dataset
    if not hasattr(dataset, "context_frame"):
        raise TypeError("CNN prediction dataset does not expose row context")
    temporary = path.with_name(path.name + ".part")
    temporary.unlink(missing_ok=True)
    offset = 0
    model.eval()
    with torch.no_grad():
        for features, targets in loader:
            predicted = model(features.to(device)).detach().cpu().numpy().reshape(-1)
            actual = targets.detach().cpu().numpy().reshape(-1)
            context = dataset.context_frame(offset, offset + len(actual))
            if len(context) != len(actual):
                raise ValueError("CNN prediction context is not aligned with model output")
            context["split"] = split
            context["y_true"] = actual
            context["y_pred"] = predicted
            context["cnn_pred"] = predicted
            context.to_csv(
                temporary,
                mode="w" if offset == 0 else "a",
                header=offset == 0,
                index=False,
                encoding="utf-8-sig" if offset == 0 else "utf-8",
            )
            offset += len(actual)
    if offset != len(dataset):
        raise ValueError("CNN prediction artifact row count does not match Test dataset")
    replace_file_atomic(temporary, path)
    return path


def train_cnn_bilstm(
    frame: pd.DataFrame,
    target_column: str,
    feature_columns: Optional[Sequence[str]] = None,
    sequence_config: Optional[SequenceConfig] = None,
    n_trials: int = 10,
    output_dir: str = "artifacts/models/cnn_bilstm",
    use_optuna: bool = True,
    use_reinforcement: bool = False,
    epochs: int = 50,
    entity_column: Optional[str] = None,
    timestamp_column: Optional[str] = None,
    optimizer_settings: OptimizationSettings | None = None,
    optimizer_trial_epochs: int = 20,
    early_stopping_patience: int = 5,
    optimizer_max_train_sequences: int | None = None,
    optimizer_max_validation_sequences: int | None = None,
    optimizer_timeout_seconds: int | None = None,
    checkpoint_store: TrainingCheckpointStore | None = None,
    checkpoint_root: str | Path | None = None,
    optimizer_storage_path: str | Path | None = None,
    optimizer_parameter_space: Mapping[str, object] | None = None,
    seed: int = 42,
) -> Dict[str, object]:
    """Train the model with Optuna and/or reinforcement learning then persist artifacts.

    All artifacts for a run are stored in a timestamped subdirectory within ``output_dir``.
    """

    cfg = sequence_config or SequenceConfig()
    random.seed(seed)
    np.random.seed(seed)
    torch.manual_seed(seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed_all(seed)
    if optimizer_parameter_space is not None and not isinstance(
        optimizer_parameter_space,
        Mapping,
    ):
        raise ValueError("optimizer_parameter_space must be an object")
    if checkpoint_store is None:
        selected_features = list(feature_columns) if feature_columns is not None else [
            column
            for column in frame.columns
            if column not in {target_column, entity_column, timestamp_column}
        ]
        resolved_checkpoint_root = (
            Path(checkpoint_root)
            if checkpoint_root is not None
            else Path(output_dir) / ".checkpoints"
        )
        checkpoint_store = TrainingCheckpointStore(
            resolved_checkpoint_root,
            model="cnn_bilstm",
            fingerprint=stable_signature(
                {
                    "checkpoint_contract": CHECKPOINT_CONTRACT,
                    "data": dataframe_signature(
                        frame,
                        [
                            *(
                                [entity_column]
                                if entity_column and entity_column in frame
                                else []
                            ),
                            *(
                                [timestamp_column]
                                if timestamp_column and timestamp_column in frame
                                else []
                            ),
                            *selected_features,
                            target_column,
                        ],
                    ),
                    "target_column": target_column,
                    "feature_columns": selected_features,
                    "sequence_config": cfg.__dict__,
                    "epochs": epochs,
                    "seed": seed,
                    "use_optuna": use_optuna,
                    "use_reinforcement": use_reinforcement,
                    "optimizer_trial_epochs": optimizer_trial_epochs,
                    "early_stopping_patience": early_stopping_patience,
                    "optimizer_max_train_sequences": optimizer_max_train_sequences,
                    "optimizer_max_validation_sequences": optimizer_max_validation_sequences,
                    "optimizer_parameter_space": dict(optimizer_parameter_space or {}),
                }
            ),
        )
    if use_optuna and optimizer_settings is None and optimizer_storage_path is not None:
        optimizer_settings = OptimizationSettings(
            enabled=True,
            study_name="cnn_pipeline_v1",
            storage_path=Path(optimizer_storage_path),
            max_trials=n_trials,
            timeout_seconds=optimizer_timeout_seconds,
            seed=seed,
            startup_trials=min(5, n_trials),
            pruner_startup_trials=min(5, n_trials),
            pruner_warmup_steps=5,
        ).scoped(checkpoint_store.fingerprint)
    run_dir = create_run_directory(Path(output_dir))

    if use_optuna:
        result = train_with_best_trial(
            frame,
            target_column,
            feature_columns=feature_columns,
            sequence_config=cfg,
            n_trials=n_trials,
            entity_column=entity_column,
            timestamp_column=timestamp_column,
            epochs=epochs,
            trial_epochs=optimizer_trial_epochs,
            early_stopping_patience=early_stopping_patience,
            maximum_train_sequences=optimizer_max_train_sequences,
            maximum_validation_sequences=optimizer_max_validation_sequences,
            settings=optimizer_settings,
            artifact_dir=run_dir,
            timeout=optimizer_timeout_seconds,
            checkpoint_store=checkpoint_store,
            optimizer_parameter_space=dict(optimizer_parameter_space or {}),
        )
        model, model_cfg = result["model"], result["model_config"]
        study = result["study"]
        checkpoint_stage = result["checkpoint_stage"]
        checkpoint_resumed = bool(result["checkpoint_resumed"])
        evaluation_loaders = result.pop("loaders")
        save_study_results(study, str(run_dir / "optuna_best.json"))
        if optimizer_settings is None:
            save_study_results(study, str(run_dir / "optimization_summary.json"))
            study.trials_dataframe().to_csv(
                run_dir / "optimization_trials.csv",
                index=False,
                encoding="utf-8-sig",
            )
    else:
        # Fallback to deterministic config
        loaders = prepare_dataset_splits(
            frame, target_column, feature_columns, cfg, entity_column, timestamp_column
        )
        train_loader, val_loader, test_loader = loaders.train, loaders.validation, loaders.test
        model_cfg = CnnBiLstmNetworkConfig(n_features=loaders.n_features)
        device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
        model = build_cnn_bilstm_network(model_cfg, device=device)
        criterion = torch.nn.MSELoss()
        optimizer = torch.optim.Adam(model.parameters(), lr=1e-3)
        checkpoint_signature = stable_signature(
            {
                "model_config": model_cfg.__dict__,
                "epochs": epochs,
                "train_sequences": len(train_loader.dataset),
                "validation_sequences": len(val_loader.dataset),
                "mode": "fixed_without_optuna",
            }
        )
        checkpoint_stage = f"final_fit_{checkpoint_signature[:20]}"
        start_epoch = 0
        checkpoint_resumed = False
        checkpoint_completed = False
        if checkpoint_store is not None:
            state = checkpoint_store.load_torch(
                checkpoint_stage,
                signature=checkpoint_signature,
                map_location=device,
            )
            if state is not None:
                model.load_state_dict(state["model_state"])
                optimizer.load_state_dict(state["optimizer_state"])
                start_epoch = int(state["next_epoch"])
                restore_rng_state(state.get("rng_state"))
                checkpoint_resumed = True
                checkpoint_completed = bool(state.get("completed", False))
        completed_epoch = start_epoch
        loop_end = start_epoch if checkpoint_completed else epochs
        for epoch in range(start_epoch, loop_end):
            train_cnn_bilstm_epoch(model, train_loader, criterion, optimizer, device)
            completed_epoch = epoch + 1
            if checkpoint_store is not None and (
                completed_epoch % checkpoint_store.cnn_every_epochs == 0
                or completed_epoch == epochs
            ):
                checkpoint_store.save_torch(
                    checkpoint_stage,
                    {
                        "model_state": model.state_dict(),
                        "optimizer_state": optimizer.state_dict(),
                        "next_epoch": completed_epoch,
                        "rng_state": capture_rng_state(),
                    },
                    signature=checkpoint_signature,
                    progress={
                        "next_epoch": completed_epoch,
                        "total_epochs": epochs,
                    },
                    completed=False,
                )
        if checkpoint_store is not None:
            checkpoint_store.save_torch(
                checkpoint_stage,
                {
                    "model_state": model.state_dict(),
                    "optimizer_state": optimizer.state_dict(),
                    "next_epoch": completed_epoch,
                    "rng_state": capture_rng_state(),
                },
                signature=checkpoint_signature,
                progress={"next_epoch": completed_epoch, "total_epochs": epochs},
                completed=True,
            )
        metrics = evaluate_cnn_bilstm_loader(model, test_loader, criterion, device)
        result = {"model": model, "model_config": model_cfg, "metrics": metrics, "best_params": {}}
        evaluation_loaders = loaders

    preprocessing_state = (
        result.get("preprocessing")
        if use_optuna
        else getattr(train_loader, "preprocessing_state", None)
    )
    temporal_split = (
        preprocessing_state.get("temporal_split")
        if isinstance(preprocessing_state, dict)
        else None
    )

    if use_reinforcement:
        train_loader, val_loader, adaptive_test_loader, _ = prepare_datasets(
            frame, target_column, feature_columns, cfg, entity_column, timestamp_column
        )
        base_checkpoint_stage = checkpoint_stage
        base_checkpoint_resumed = checkpoint_resumed
        adaptive_config = BanditConfig(actions=[1e-4, 5e-4, 1e-3, 5e-3])
        adaptive_signature = stable_signature(
            {
                "model_config": model_cfg.__dict__,
                "bandit_config": adaptive_config.__dict__,
                "epochs": 30,
                "train_sequences": len(train_loader.dataset),
                "validation_sequences": len(val_loader.dataset),
                "upstream_stage": base_checkpoint_stage,
            }
        )
        checkpoint_stage = f"adaptive_fit_{adaptive_signature[:20]}"
        reinforcement = run_adaptive_training(
            model_cfg,
            train_loader,
            val_loader,
            epochs=30,
            bandit_cfg=adaptive_config,
            checkpoint_store=checkpoint_store,
            checkpoint_stage=checkpoint_stage,
            checkpoint_signature=adaptive_signature,
            initial_model=model,
        )
        model = reinforcement.model
        checkpoint_resumed = reinforcement.checkpoint_resumed
        adaptive_device = next(model.parameters()).device
        result["metrics"] = evaluate_cnn_bilstm_loader(
            model,
            adaptive_test_loader,
            torch.nn.MSELoss(),
            adaptive_device,
        )
        with open(run_dir / "reinforcement_history.json", "w", encoding="utf-8") as f:
            json.dump(reinforcement.history, f, indent=2)
        with open(run_dir / "bandit_q_values.json", "w", encoding="utf-8") as f:
            json.dump(reinforcement.q_values, f, indent=2)

    checkpoint_path = run_dir / "cnn_bilstm.pt"
    temporary_checkpoint = checkpoint_path.with_name(
        f"{checkpoint_path.stem}.tmp{checkpoint_path.suffix}"
    )
    torch.save(
        {
            "model_state": model.state_dict(),
            "config": model_cfg.__dict__,
            "feature_columns": list(feature_columns or []),
            "preprocessing": preprocessing_state,
        },
        temporary_checkpoint,
    )
    replace_file_atomic(temporary_checkpoint, checkpoint_path)
    with open(run_dir / "metrics.json", "w", encoding="utf-8") as f:
        json.dump(result.get("metrics", {}), f, indent=2)
    with open(run_dir / "best_params.json", "w", encoding="utf-8") as f:
        json.dump(result.get("best_params", {}), f, indent=2)

    prediction_device = next(model.parameters()).device
    prediction_paths = {
        split: _write_prediction_artifact(
            model,
            getattr(evaluation_loaders, split),
            run_dir / f"{split}_predictions.csv",
            split=split,
            device=prediction_device,
        )
        for split in ("validation", "calibration", "test")
    }

    optimizer_artifact = (
        {
            "enabled": True,
            "selection_data": "validation_only",
            "objective_metric": "validation_mae",
            "best_validation_mae": float(study.best_value),
            "best_params": dict(study.best_params),
            "summary_path": str(run_dir / "optimization_summary.json"),
            "trials_path": str(run_dir / "optimization_trials.csv"),
            "test_usage": "none",
        }
        if use_optuna
        else {
            "enabled": False,
            "reason": "disabled_by_config",
        }
    )

    checkpoint_artifact = (
        {
            **checkpoint_store.describe(),
            "stage": checkpoint_stage,
            "upstream_stage": base_checkpoint_stage if use_reinforcement else None,
            "upstream_resumed": (
                base_checkpoint_resumed if use_reinforcement else None
            ),
            "resumed": checkpoint_resumed,
            "retained_for_idempotent_resume": True,
        }
        if checkpoint_store is not None
        else {
            "enabled": False,
            "resume": False,
            "reason": "checkpoint_store_not_configured",
        }
    )
    return {
        "model": model,
        "config": model_cfg,
        "metrics": result.get("metrics", {}),
        "output_dir": str(run_dir),
        "checkpoint_path": str(checkpoint_path),
        "validation_predictions": str(prediction_paths["validation"]),
        "calibration_predictions": str(prediction_paths["calibration"]),
        "test_predictions": str(prediction_paths["test"]),
        "temporal_split": temporal_split,
        "optimizer": optimizer_artifact,
        "checkpoint": checkpoint_artifact,
    }
