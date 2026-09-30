"""Leakage-safe model feature construction."""
from importlib import import_module

_EXPORTS = {'SELECTED_MODEL_FEATURES': ('solar_forecast.features.history_features', 'SELECTED_MODEL_FEATURES'), 'LeakageSafeFeatureEngineer': ('solar_forecast.features.history_features', 'LeakageSafeFeatureEngineer'), 'LegacyModelDatasetBuilder': ('solar_forecast.datasets.model_dataset_builder', 'LegacyModelDatasetBuilder'), 'ModelDatasetResult': ('solar_forecast.datasets.model_dataset_builder', 'ModelDatasetResult'), 'NationwideModelDatasetBuilder': ('solar_forecast.datasets.model_dataset_builder', 'NationwideModelDatasetBuilder'), 'KmaAsosNormalizer': ('solar_forecast.features.asos_features', 'KmaAsosNormalizer')}
__all__ = list(_EXPORTS)

def __getattr__(name: str):
    """Load a public symbol only when its owning feature is requested."""
    if name not in _EXPORTS:
        raise AttributeError(name)
    module, symbol = _EXPORTS[name]
    value = getattr(import_module(module), symbol)
    globals()[name] = value
    return value
