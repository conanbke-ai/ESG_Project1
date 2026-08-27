"""Anomaly-signal policies that avoid unsupported equipment-failure claims."""

from .policy import ALLOWED_INFLUENCE_FACTORS, INTERPRETATION_LIMIT, validate_influence_factor

__all__ = ["ALLOWED_INFLUENCE_FACTORS", "INTERPRETATION_LIMIT", "validate_influence_factor"]
