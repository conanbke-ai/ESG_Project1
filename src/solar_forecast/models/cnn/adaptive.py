"""Simple reinforcement-learning utilities for adaptive training."""
from __future__ import annotations

import random
from dataclasses import dataclass, field
from typing import Dict, List, Sequence

import numpy as np
import torch
import torch.nn as nn

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


def run_adaptive_training(
    model_cfg: ModelConfig,
    train_loader,
    val_loader,
    epochs: int = 30,
    bandit_cfg: BanditConfig | None = None,
    device: torch.device | None = None,
) -> ReinforcementResult:
    """Train the model while a bandit adapts the learning rate."""

    device = device or torch.device("cuda" if torch.cuda.is_available() else "cpu")
    bandit_cfg = bandit_cfg or BanditConfig(actions=[1e-4, 5e-4, 1e-3, 5e-3])
    state = BanditState()
    state.initialize(bandit_cfg.actions)

    model = build_model(model_cfg, device=device)
    criterion = nn.MSELoss()
    optimizer = torch.optim.Adam(model.parameters(), lr=bandit_cfg.actions[0])

    history: List[Dict[str, float]] = []
    best_val = float("inf")
    best_state = None

    for epoch in range(epochs):
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
            best_state = model.state_dict()

    if best_state is not None:
        model.load_state_dict(best_state)

    return ReinforcementResult(model=model, history=history, q_values=state.q_values)
