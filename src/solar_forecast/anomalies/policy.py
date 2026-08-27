from __future__ import annotations


ALLOWED_INFLUENCE_FACTORS = frozenset(
    {
        "weather_impact",
        "lower_than_expected",
        "higher_than_expected",
        "rapid_output_change",
        "generation_stop_suspected",
        "data_quality_issue",
        "plant_specific_unknown_factor",
        "unknown_external_factor",
    }
)

INTERPRETATION_LIMIT = (
    "공개 데이터만으로 설비 고장·정비·출력제어 여부를 확인할 수 없습니다."
)


def validate_influence_factor(value: str) -> str:
    """Reject unsupported root-cause labels before they reach reports or alerts."""
    if value not in ALLOWED_INFLUENCE_FACTORS:
        allowed = ", ".join(sorted(ALLOWED_INFLUENCE_FACTORS))
        raise ValueError(f"Unsupported influence factor: {value}. Allowed: {allowed}")
    return value
