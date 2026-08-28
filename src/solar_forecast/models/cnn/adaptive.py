"""Simple reinforcement-learning utilities for adaptive training."""
from __future__ import annotations

from copy import deepcopy
import random
from dataclasses import dataclass, field
from typing import Dict, List, Sequence

import numpy as np
import torch
import torch.nn as nn

from solar_forecast.models.checkpointing import (
    TrainingCheckpointStore,
    capture_rng_state,
    restore_rng_state,
)

from .model import CNNBiLSTM, ModelConfig, build_model
from .optimization import _evaluate, _train_one_epoch


@dataclass
class BanditConfig:
    """Configuration for the epsilon-greedy multi-armed bandit."""

    actions: Sequence[float]
    epsilon: float = 0.15
    alpha: float = 0.2
    reward_scale: float = 10.0


@dataclass
class BanditState:
    q_values: Dict[float, float] = field(default_factory=dict)

    def initialize(self, actions: Sequence[float]):
        for action in actions:
            self.q_values.setdefault(float(action), 0.0)

    def select(self, actions: Sequence[float], epsilon: float) -> float:
        if random.random() < epsilon:
            return float(random.choice(actions))
        return float(max(actions, key=lambda a: self.q_values.get(float(a), 0.0)))

    def update(self, action: float, reward: float, alpha: float):
        current = self.q_values.get(float(action), 0.0)
        self.q_values[float(action)] = current + alpha * (reward - current)


@dataclass
class ReinforcementResult:
    model: CNNBiLSTM
    history: List[Dict[str, float]]
    q_values: Dict[float, float]
    checkpoint_resumed: bool = False


def run_adaptive_training(
    model_cfg: ModelConfig,
    train_loader,
    val_loader,
    epochs: int = 30,
    bandit_cfg: BanditConfig | None = None,
    device: torch.device | None = None,
    checkpoint_store: TrainingCheckpointStore | None = None,
    checkpoint_stage: str = "adaptive_fit",
    checkpoint_signature: str = "adaptive_fit_v1",
    initial_model: CNNBiLSTM | None = None,
) -> ReinforcementResult:
    """Train the model while a bandit adapts the learning rate."""

    device = device or torch.device("cuda" if torch.cuda.is_available() else "cpu")
    bandit_cfg = bandit_cfg or BanditConfig(actions=[1e-4, 5e-4, 1e-3, 5e-3])
    state = BanditState()
    state.initialize(bandit_cfg.actions)

    model = initial_model.to(device) if initial_model is not None else build_model(
        model_cfg,
        device=device,
    )
    criterion = nn.MSELoss()
    optimizer = torch.optim.Adam(model.parameters(), lr=bandit_cfg.actions[0])

    history: List[Dict[str, float]] = []
    best_val = float("inf")
    best_state = None
    start_epoch = 0
    checkpoint_resumed = False
    checkpoint_completed = False

    if checkpoint_store is not None:
        checkpoint = checkpoint_store.load_torch(
            checkpoint_stage,
            signature=checkpoint_signature,
            map_location=device,
        )
        if checkpoint is not None:
            model.load_state_dict(checkpoint["model_state"])
            optimizer.load_state_dict(checkpoint["optimizer_state"])
            state.q_values = {
                float(action): float(value)
                for action, value in checkpoint["q_values"].items()
            }
            history = list(checkpoint.get("history", []))
            best_val = float(checkpoint.get("best_val", float("inf")))
            best_state = checkpoint.get("best_state")
            start_epoch = int(checkpoint["next_epoch"])
            restore_rng_state(checkpoint.get("rng_state"))
            checkpoint_resumed = True
            checkpoint_completed = bool(checkpoint.get("completed", False))

    loop_end = start_epoch if checkpoint_completed else epochs
    for epoch in range(start_epoch, loop_end):
        lr_choice = state.select(bandit_cfg.actions, bandit_cfg.epsilon)
        for g in optimizer.param_groups:
            g["lr"] = lr_choice

        train_loss = _train_one_epoch(model, train_loader, criterion, optimizer, device)
        metrics = _evaluate(model, val_loader, criterion, device)
        reward = -metrics["loss"] * bandit_cfg.reward_scale
        state.update(lr_choice, reward, bandit_cfg.alpha)

        history.append(
            {
                "epoch": epoch + 1,
                "lr": lr_choice,
                "train_loss": train_loss,
                **metrics,
            }
        )

        if metrics["loss"] < best_val:
            best_val = metrics["loss"]
            best_state = deepcopy(model.state_dict())

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
                    "q_values": state.q_values,
                    "history": history,
                    "best_val": best_val,
                    "best_state": best_state,
                    "next_epoch": completed_epoch,
                    "rng_state": capture_rng_state(),
                },
                signature=checkpoint_signature,
                progress={"next_epoch": completed_epoch, "total_epochs": epochs},
                completed=False,
            )

    if checkpoint_store is not None:
        checkpoint_store.save_torch(
            checkpoint_stage,
            {
                "model_state": model.state_dict(),
                "optimizer_state": optimizer.state_dict(),
                "q_values": state.q_values,
                "history": history,
                "best_val": best_val,
                "best_state": best_state,
                "next_epoch": len(history),
                "rng_state": capture_rng_state(),
            },
            signature=checkpoint_signature,
            progress={"next_epoch": len(history), "total_epochs": epochs},
            completed=True,
        )

    if best_state is not None:
        model.load_state_dict(best_state)

    return ReinforcementResult(
        model=model,
        history=history,
        q_values=state.q_values,
        checkpoint_resumed=checkpoint_resumed,
    )
