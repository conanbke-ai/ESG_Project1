import json

from solar_forecast.infrastructure.error_report import write_error_report


def test_error_report_contains_stage_context_and_traceback(tmp_path):
    try:
        raise ValueError("bad input")
    except ValueError as exc:
        path = write_error_report(tmp_path, exc, stage="preprocessing", context={"target": "generation"})

    content = path.read_text(encoding="utf-8")
    assert path.name == "error.log"
    assert '"stage": "preprocessing"' in content
    assert '"target": "generation"' in content
    assert "ValueError: bad input" in content
