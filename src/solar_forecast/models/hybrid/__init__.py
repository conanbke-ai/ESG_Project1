"""Reproducible controlled, optimized, and hybrid experiments."""
from importlib import import_module

_EXPORTS = {'DynamicGateConfig': ('solar_forecast.models.hybrid.dynamic_gate', 'DynamicGateConfig'), 'ExplainableDynamicGate': ('solar_forecast.models.hybrid.dynamic_gate', 'ExplainableDynamicGate'), 'fit_dynamic_gate': ('solar_forecast.models.hybrid.dynamic_gate', 'fit_dynamic_gate'), 'fit_region_blend': ('solar_forecast.models.hybrid.dynamic_gate', 'fit_region_blend'), 'predict_dynamic_hybrid': ('solar_forecast.models.hybrid.dynamic_gate', 'predict_dynamic_hybrid'), 'predict_region_blend': ('solar_forecast.models.hybrid.dynamic_gate', 'predict_region_blend'), 'aggregate_metrics': ('solar_forecast.evaluation.regression_metrics', 'aggregate_metrics'), 'calculate_metrics': ('solar_forecast.evaluation.regression_metrics', 'calculate_metrics')}
__all__ = list(_EXPORTS)

def __getattr__(name: str):
    """Load a public symbol only when its owning feature is requested."""
    if name not in _EXPORTS:
        raise AttributeError(name)
    module, symbol = _EXPORTS[name]
    value = getattr(import_module(module), symbol)
    globals()[name] = value
    return value
