"""Validate fixed-lead future-weather archives before model training."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

import pandas as pd

from solar_forecast.features.future_weather import (
    FUTURE_WEATHER_CONTRACT,
    future_weather_feature_columns,
    read_future_weather,
)


RANGES = {
    "future_temperature_c": (-60.0, 60.0),
    "future_humidity_pct": (0.0, 100.0),
    "future_precipitation_mm": (0.0, None),
    "future_total_cloud_cover_tenths": (0.0, 10.0),
    "future_wind_speed_mps": (0.0, None),
    "future_solar_irradiance_mj_m2": (0.0, 6.0),
    "future_sunshine_hours": (0.0, 1.0),
    "future_dni_mj_m2": (0.0, 6.0),
    "future_dhi_mj_m2": (0.0, 6.0),
}


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        default="file/forecast_weather",
    )
    parser.add_argument(
        "--output",
        default="artifacts/verification/future_weather_preflight.json",
    )
    return parser


def _manifest(path: Path) -> dict:
    manifest_path = path.with_name(
        path.name.replace(".csv.gz", ".manifest.json")
    )
    if not manifest_path.is_file():
        raise FileNotFoundError(
            f"Future-weather manifest is missing: {manifest_path}"
        )
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    if payload.get("contract") != FUTURE_WEATHER_CONTRACT:
        raise ValueError(
            f"{manifest_path}: expected contract {FUTURE_WEATHER_CONTRACT}, "
            f"got {payload.get('contract')!r}"
        )
    return payload


def _feature_stats(frame: pd.DataFrame, features: tuple[str, ...]) -> dict:
    result = {}
    for column in features:
        values = pd.to_numeric(frame[column], errors="coerce")
        observed = values.dropna()
        low, high = RANGES[column]
        invalid = values.notna() & (
            (values < low)
            | ((values > high) if high is not None else False)
        )
        result[column] = {
            "observed_rows": int(values.notna().sum()),
            "missing_rows": int(values.isna().sum()),
            "missing_fraction": float(values.isna().mean()),
            "min": float(observed.min()) if not observed.empty else None,
            "max": float(observed.max()) if not observed.empty else None,
            "invalid_rows": int(invalid.sum()),
        }
    return result


def validate(root: Path) -> dict:
    reports = {}
    key_sets = {}
    for horizon in (24, 72):
        path = root / f"open_meteo_jma_msm_{horizon}h.csv.gz"
        manifest = _manifest(path)
        frame = read_future_weather(
            path,
            horizon_hours=horizon,
            feature_profile="aligned_core_plus_components",
        )
        core = future_weather_feature_columns("aligned_core")
        extended = future_weather_feature_columns(
            "aligned_core_plus_components"
        )
        feature_stats = _feature_stats(frame, extended)
        invalid_features = [
            name
            for name, stats in feature_stats.items()
            if stats["invalid_rows"] > 0
        ]
        keys = pd.MultiIndex.from_frame(frame[["plant_id", "timestamp"]])
        key_sets[horizon] = keys
        reports[str(horizon)] = {
            "manifest_contract": manifest["contract"],
            "normalization_contract": manifest.get("normalization_contract"),
            "rows": int(len(frame)),
            "plants": int(frame["plant_id"].nunique()),
            "start": frame["timestamp"].min().isoformat(),
            "end": frame["timestamp"].max().isoformat(),
            "coordinate_sources": (
                frame["coordinate_source"].value_counts().sort_index().to_dict()
            ),
            "aligned_core_features": list(core),
            "extended_features": list(extended),
            "feature_stats": feature_stats,
            "invalid_features": invalid_features,
            "fixed_lead_verified": True,
        }

    same_keys = key_sets[24].equals(key_sets[72])
    invalid = {
        horizon: report["invalid_features"]
        for horizon, report in reports.items()
        if report["invalid_features"]
    }
    report = {
        "contract": "solar-future-weather-preflight.v1",
        "status": "PASS" if same_keys and not invalid else "FAIL",
        "future_weather_contract": FUTURE_WEATHER_CONTRACT,
        "horizons": reports,
        "same_plant_timestamp_keys_24h_72h": same_keys,
        "invalid_features": invalid,
        "historical_alignment": {
            "temperature": "Celsius_to_Celsius",
            "humidity": "percent_to_percent",
            "precipitation": "mm_to_mm",
            "total_cloud_cover": "forecast_percent_divided_by_10_to_ASOS_tenths",
            "wind_speed": "mps_to_mps",
            "solar_irradiance": "forecast_hourly_mean_W_m2_times_0.0036_to_MJ_m2",
            "sunshine": "forecast_seconds_divided_by_3600_to_hours",
            "dni_dhi": "ablation_only_no_direct_current_ASOS_counterpart",
        },
    }
    return report


def main() -> None:
    args = _parser().parse_args()
    root = Path(args.root)
    output = Path(args.output)
    report = validate(root)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if report["status"] != "PASS":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
