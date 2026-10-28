"""Canonical Solar plant-data preprocessing boundary.

This package owns preprocessing orchestration and training-population eligibility.
Low-level storage/registry implementations remain in datasets/, quality rules in
quality/, and model performance evaluation in evaluation/.
"""
from importlib import import_module

_EXPORTS = {
    "DataPreparationResult": ("solar_forecast.preprocessing.preparation_service", "DataPreparationResult"),
    "DataPreparationService": ("solar_forecast.preprocessing.preparation_service", "DataPreparationService"),
    "audit_training_eligibility": ("solar_forecast.preprocessing.training_eligibility", "audit_training_eligibility"),
    "run_training_eligibility_audit": ("solar_forecast.preprocessing.training_eligibility", "run_training_eligibility_audit"),
    "classify_training_population": ("solar_forecast.preprocessing.training_admission", "classify_training_population"),
    "validate_admission_policy": ("solar_forecast.preprocessing.training_admission", "validate_admission_policy"),
}
__all__ = list(_EXPORTS)

def __getattr__(name: str):
    if name not in _EXPORTS:
        raise AttributeError(name)
    module, symbol = _EXPORTS[name]
    value = getattr(import_module(module), symbol)
    globals()[name] = value
    return value
