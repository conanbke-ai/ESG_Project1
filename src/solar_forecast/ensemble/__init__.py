"""Reproducible controlled, optimized, and hybrid experiments."""

from .dynamic_gate import (
    DynamicGateConfig, ExplainableDynamicGate, fit_dynamic_gate, fit_region_blend,
    predict_dynamic_hybrid, predict_region_blend,
)
from .metrics import aggregate_metrics, calculate_metrics

__all__ = [
    "DynamicGateConfig", "ExplainableDynamicGate", "fit_dynamic_gate",
    "predict_dynamic_hybrid", "fit_region_blend", "predict_region_blend",
    "aggregate_metrics", "calculate_metrics",
]
