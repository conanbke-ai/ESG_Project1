"""xgboost model public training API."""
from importlib import import_module

_EXPORTS = {'XGBoostTrainer': ('solar_forecast.models.xgboost.trainer', 'XGBoostTrainer'), 'train': ('solar_forecast.models.xgboost.trainer', 'train')}
__all__ = list(_EXPORTS)

def __getattr__(name: str):
    """Load a public symbol only when its owning feature is requested."""
    if name not in _EXPORTS:
        raise AttributeError(name)
    module, symbol = _EXPORTS[name]
    value = getattr(import_module(module), symbol)
    globals()[name] = value
    return value
