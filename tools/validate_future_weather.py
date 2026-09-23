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
    parser.add_argument(
        "--minimum-coverage",
        type=float,
        default=0.98,
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


def _all_missing_breakdown(
    frame: pd.DataFrame,
    features: tuple[str, ...],
) -> dict[str, object]:
    missing = frame[list(features)].isna().all(axis=1)
    affected = frame.loc[missing, ["plant_id", "timestamp"]].copy()
    return {
        "rows": int(missing.sum()),
        "fraction": float(missing.mean()) if len(frame) else 0.0,
        "plants": int(affected["plant_id"].nunique()) if not affected.empty else 0,
        "plant_rows": (
            affected["plant_id"].value_counts().sort_index().to_dict()
            if not affected.empty
            else {}
        ),
        "start": (
            affected["timestamp"].min().isoformat()
            if not affected.empty
            else None
        ),
        "end": (
            affected["timestamp"].max().isoformat()
            if not affected.empty
            else None
        ),
    }


def validate(root: Path, *, minimum_coverage: float = 0.98) -> dict:
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
        entirely_missing_features = [
            name
            for name, stats in feature_stats.items()
            if stats["observed_rows"] == 0
        ]
        core_missing = _all_missing_breakdown(frame, core)
        extended_missing = _all_missing_breakdown(frame, extended)
        core_coverage = 1.0 - float(core_missing["fraction"])
        extended_coverage = 1.0 - float(extended_missing["fraction"])
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
            "entirely_missing_features": entirely_missing_features,
            "aligned_core_all_missing": core_missing,
            "extended_all_missing": extended_missing,
            "aligned_core_coverage": core_coverage,
            "extended_coverage": extended_coverage,
            "minimum_required_coverage": float(minimum_coverage),
            "coverage_passed": (
                core_coverage >= minimum_coverage
                and extended_coverage >= minimum_coverage
            ),
            "fixed_lead_verified": True,
        }

    same_keys = key_sets[24].equals(key_sets[72])
    invalid = {
        horizon: report["invalid_features"]
        for horizon, report in reports.items()
        if report["invalid_features"]
    }
    entirely_missing = {
        horizon: report["entirely_missing_features"]
        for horizon, report in reports.items()
        if report["entirely_missing_features"]
    }
    coverage_failures = {
        horizon: {
            "aligned_core_coverage": report["aligned_core_coverage"],
            "extended_coverage": report["extended_coverage"],
        }
        for horizon, report in reports.items()
        if not report["coverage_passed"]
    }
    passed = (
        same_keys
        and not invalid
        and not entirely_missing
        and not coverage_failures
    )
    report = {
        "contract": "solar-future-weather-preflight.v1",
        "status": "PASS" if passed else "FAIL",
        "future_weather_contract": FUTURE_WEATHER_CONTRACT,
        "horizons": reports,
        "same_plant_timestamp_keys_24h_72h": same_keys,
        "invalid_features": invalid,
        "entirely_missing_features": entirely_missing,
        "coverage_failures": coverage_failures,
        "minimum_required_coverage": float(minimum_coverage),
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
    if not 0 < args.minimum_coverage <= 1:
        raise ValueError("--minimum-coverage must be in (0, 1]")
    report = validate(
        root,
        minimum_coverage=float(args.minimum_coverage),
    )
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
