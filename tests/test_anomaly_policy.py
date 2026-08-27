import pytest

from solar_forecast.anomalies.policy import validate_influence_factor


def test_supported_influence_factor_is_accepted() -> None:
    assert validate_influence_factor("weather_impact") == "weather_impact"


def test_equipment_failure_claim_is_rejected() -> None:
    with pytest.raises(ValueError):
        validate_influence_factor("equipment_failure")
