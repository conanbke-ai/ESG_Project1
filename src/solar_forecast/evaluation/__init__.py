"""Leakage-safe evaluation services."""
from importlib import import_module

_EXPORTS = {'FeatureAblationResult': ('solar_forecast.evaluation.feature_ablation', 'FeatureAblationResult'), 'FeatureAblationService': ('solar_forecast.evaluation.feature_ablation', 'FeatureAblationService'), 'TemporalBoundaries': ('solar_forecast.evaluation.temporal_split', 'TemporalBoundaries'), 'TemporalFrameSplits': ('solar_forecast.evaluation.temporal_split', 'TemporalFrameSplits'), 'TemporalSplitConfig': ('solar_forecast.evaluation.temporal_split', 'TemporalSplitConfig'), 'TemporalSplitter': ('solar_forecast.evaluation.temporal_split', 'TemporalSplitter')}
__all__ = list(_EXPORTS)

def __getattr__(name: str):
    """Load a public symbol only when its owning feature is requested."""
    if name not in _EXPORTS:
        raise AttributeError(name)
    module, symbol = _EXPORTS[name]
    value = getattr(import_module(module), symbol)
    globals()[name] = value
    return value
