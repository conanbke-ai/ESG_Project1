"""cnn_bilstm model public training API."""
from importlib import import_module

_EXPORTS = {'CnnBiLstmTrainer': ('solar_forecast.models.cnn_bilstm.trainer', 'CnnBiLstmTrainer'), 'train': ('solar_forecast.models.cnn_bilstm.trainer', 'train'), 'SequenceConfig': ('solar_forecast.models.cnn_bilstm.sequence_config', 'SequenceConfig')}
__all__ = list(_EXPORTS)

def __getattr__(name: str):
    """Load a public symbol only when its owning feature is requested."""
    if name not in _EXPORTS:
        raise AttributeError(name)
    module, symbol = _EXPORTS[name]
    value = getattr(import_module(module), symbol)
    globals()[name] = value
    return value
