"""Verify source and hourly coverage evidence in the archive preflight."""

import hashlib
import json

import pandas as pd
import pytest

from solar_forecast.features.future_weather import (
    FUTURE_WEATHER_CONTRACT,
    FUTURE_WEATHER_METEOROLOGY_FEATURES,
)
from tools.validate_future_weather import validate


@pytest.fixture
def weather_preflight(tmp_path):
    archive = tmp_path / "weather.csv.gz"
    manifest = tmp_path / "weather.manifest.json"
    config = tmp_path / "experiment.json"
    timestamps = pd.to_datetime(["2025-06-01 00:00", "2025-06-01 02:00", "2025-06-01 00:00"])
    frame = pd.DataFrame(
        {
            "plant_id": ["plant-a", "plant-a", "plant-b"],
            "timestamp": timestamps,
            "forecast_origin": timestamps - pd.Timedelta(hours=72),
            "horizon_hours": 72,
            "weather_query_latitude": 35.0,
            "weather_query_longitude": 129.0,
            "grid_latitude": 35.0,
            "grid_longitude": 129.0,
            "coordinate_source": "plant_registry_coordinates",
            "cell_selection": "nearest",
            "forecast_model": "jma_gsm",
            "forecast_source": "open_meteo_previous_runs",
        }
    )
    frame[list(FUTURE_WEATHER_METEOROLOGY_FEATURES)] = 1.0
    manifest.write_text(json.dumps({"contract": FUTURE_WEATHER_CONTRACT, "model": "jma_gsm"}))
    config.write_text(json.dumps({"future_weather": {
        "enabled_horizons": [72],
        "by_horizon": {"72": {
            "source": str(archive),
            "source_model": "jma_gsm",
            "feature_profiles": ["aligned_meteorology_core"],
        }},
    }}))
    return archive, manifest, config, frame


def test_archive_key_gaps_and_missing_feature_rows_are_separate(weather_preflight):
    archive, _, config, frame = weather_preflight
    frame.loc[0, list(FUTURE_WEATHER_METEOROLOGY_FEATURES)] = float("nan")
    frame.to_csv(archive, index=False)

    report = validate(config, root=archive.parent, minimum_coverage=0.6)
    horizon = report["horizons"]["72"]

    assert report["status"] == "PASS"
    assert horizon["source_model_matches"]
    assert horizon["as_of_evidence"]["status"] == "UNVERIFIED"
    assert horizon["archive_hourly_keys"] == {
        "invalid_hour_rows": 0,
        "missing_plant_id_rows": 0,
        "expected_in_observed_span": 6,
        "missing_in_observed_span": 3,
        "missing_hours_by_plant": {"plant-a": 1, "plant-b": 2},
    }
    profile = horizon["profiles"]["aligned_meteorology_core"]
    assert profile["all_missing"]["rows"] == 1
    assert profile["coverage"] == pytest.approx(2 / 3)


@pytest.mark.parametrize("wrong_manifest", [False, True])
def test_preflight_checks_configured_model_against_manifest_and_rows(weather_preflight, wrong_manifest):
    archive, manifest, config, frame = weather_preflight
    if wrong_manifest:
        manifest.write_text(json.dumps({"contract": FUTURE_WEATHER_CONTRACT, "model": "jma_msm"}))
    else:
        frame.loc[0, "forecast_model"] = "jma_msm"
    frame.to_csv(archive, index=False)

    report = validate(config, root=archive.parent)

    assert report["status"] == "FAIL"
    assert report["archive_contract_failures"] == ["72"]
    assert not report["horizons"]["72"]["source_model_matches"]


def test_preflight_rejects_off_hour_keys_and_infinite_weather(weather_preflight):
    archive, _, config, frame = weather_preflight
    for column in ("timestamp", "forecast_origin"):
        frame.loc[0, column] += pd.Timedelta(minutes=15)
    frame.loc[1, "future_precipitation_mm"] = float("inf")
    frame.loc[2, "plant_id"] = None
    frame.to_csv(archive, index=False)

    report = validate(config, root=archive.parent)

    assert report["status"] == "FAIL"
    assert report["archive_contract_failures"] == ["72"]
    horizon = report["horizons"]["72"]
    assert horizon["archive_hourly_keys"]["invalid_hour_rows"] == 1
    assert horizon["archive_hourly_keys"]["missing_plant_id_rows"] == 1
    assert horizon["invalid_features"] == ["future_precipitation_mm"]
    assert horizon["feature_stats"]["future_precipitation_mm"]["max"] == 1.0
    json.dumps(report, allow_nan=False)


def test_preflight_reports_timezone_representation_mismatch(weather_preflight):
    archive, _, config, frame = weather_preflight
    for column in ("timestamp", "forecast_origin"):
        frame[column] = frame[column].dt.tz_localize("UTC")
    frame.to_csv(archive, index=False)
    settings = json.loads(config.read_text())
    settings["future_weather"]["by_horizon"]["72"]["usable_from"] = "2025-06-01"
    config.write_text(json.dumps(settings))

    with pytest.raises(ValueError, match="must both have a timezone"):
        validate(config, root=archive.parent)


@pytest.mark.parametrize(
    ("available_at", "archive_digest_matches", "status", "error"),
    [
        ("2025-05-28T14:57:00Z", True, "PARTIAL", None),
        ("2025-05-28T15:01:00Z", True, "FAIL", "available_after_forecast_origin"),
        ("2025-05-28T14:49:00Z", True, "FAIL", "issued_after_available"),
        ("2025-05-28T14:57:00Z", False, "FAIL", "archive_hash_mismatch"),
    ],
)
def test_timing_evidence_is_matched_to_archive_and_cutoff(
    weather_preflight, available_at, archive_digest_matches, status, error
):
    archive, _, config, frame = weather_preflight
    frame.to_csv(archive, index=False)
    digest = hashlib.sha256(archive.read_bytes()).hexdigest()
    timing = archive.parent / "timestamps.csv"
    pd.DataFrame([{
        "plant_id": "plant-a",
        "horizon_hours": 72,
        "forecast_origin": "2025-05-29T00:00:00+09:00",
        "issued_at": "2025-05-28T14:50:00Z",
        "available_at": available_at,
        "valid_at": "2025-05-31T15:00:00Z",
        "archive_sha256": digest if archive_digest_matches else "0" * 64,
        "evidence_ref": "provider-run-receipt/1",
    }]).to_csv(timing, index=False)

    report = validate(config, root=archive.parent, timing_evidence_path=timing)

    evidence = report["horizons"]["72"]["as_of_evidence"]
    assert evidence["status"] == status
    assert evidence["rows_checked"] == 1
    assert evidence["row_errors"]["no_usable_archive_key"] == 0
    if error is not None:
        assert evidence["rows_verified"] == 0
        assert evidence["row_errors"][error] == 1
        assert report["status"] == "FAIL"
    else:
        assert evidence["rows_verified"] == 1
        assert report["status"] == "PASS"
