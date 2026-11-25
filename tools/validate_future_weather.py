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
        "--config",
        default="config/experiments/optimized.json",
        help="Experiment config containing horizon-specific future-weather sources.",
    )
    parser.add_argument(
        "--root",
        default="file/forecast_weather",
        help="Fallback root for relative archive paths.",
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


def validate(
    config_path: Path,
    *,
    root: Path,
    minimum_coverage: float = 0.98,
) -> dict:
    values = json.loads(config_path.read_text(encoding="utf-8"))
    future_config = values.get("future_weather") or {}
    enabled_horizons = list(future_config.get("enabled_horizons", []))
    by_horizon = future_config.get("by_horizon") or {}
    if not enabled_horizons:
        raise ValueError("Experiment config has no future-weather horizons")

    reports = {}
    key_sets = {}
    for horizon in enabled_horizons:
        spec = by_horizon.get(str(horizon), by_horizon.get(horizon))
        if not isinstance(spec, dict):
            raise ValueError(
                f"Missing future-weather configuration for {horizon}h"
            )
        source = Path(str(spec["source"]))
        if not source.is_absolute():
            source = PROJECT_ROOT / source
        manifest = _manifest(source)
        profiles = tuple(map(str, spec.get("feature_profiles", [])))
        if not profiles:
            raise ValueError(f"{horizon}h has no feature profiles")

        broadest_profile = (
            "aligned_core_plus_components"
            if "aligned_core_plus_components" in profiles
            else profiles[-1]
        )
        frame = read_future_weather(
            source,
            horizon_hours=int(horizon),
            feature_profile=broadest_profile,
        )
        usable_from = spec.get("usable_from")
        if usable_from:
            usable_timestamp = pd.Timestamp(str(usable_from))
            cohort = frame.loc[frame["timestamp"].ge(usable_timestamp)].copy()
        else:
            cohort = frame.copy()
        if cohort.empty:
            raise ValueError(f"{horizon}h usable cohort is empty")

        profile_reports = {}
        all_selected = future_weather_feature_columns(broadest_profile)
        feature_stats = _feature_stats(cohort, all_selected)
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
        for profile in profiles:
            features = future_weather_feature_columns(profile)
            missing = _all_missing_breakdown(cohort, features)
            coverage = 1.0 - float(missing["fraction"])
            profile_reports[profile] = {
                "features": list(features),
                "all_missing": missing,
                "coverage": coverage,
                "minimum_required_coverage": float(minimum_coverage),
                "coverage_passed": coverage >= minimum_coverage,
            }

        keys = pd.MultiIndex.from_frame(cohort[["plant_id", "timestamp"]])
        key_sets[int(horizon)] = keys
        reports[str(horizon)] = {
            "source": str(source),
            "source_model": spec.get("source_model"),
            "manifest_contract": manifest["contract"],
            "normalization_contract": manifest.get("normalization_contract"),
            "rows_total": int(len(frame)),
            "rows_usable": int(len(cohort)),
            "plants": int(cohort["plant_id"].nunique()),
            "start_total": frame["timestamp"].min().isoformat(),
            "end_total": frame["timestamp"].max().isoformat(),
            "usable_from": usable_from,
            "start_usable": cohort["timestamp"].min().isoformat(),
            "end_usable": cohort["timestamp"].max().isoformat(),
            "coordinate_sources": (
                cohort["coordinate_source"].value_counts().sort_index().to_dict()
            ),
            "profiles": profile_reports,
            "feature_stats": feature_stats,
            "invalid_features": invalid_features,
            "entirely_missing_features": entirely_missing_features,
            "fixed_lead_verified": True,
        }

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
            profile: details["coverage"]
            for profile, details in report["profiles"].items()
            if not details["coverage_passed"]
        }
        for horizon, report in reports.items()
        if any(
            not details["coverage_passed"]
            for details in report["profiles"].values()
        )
    }
    coverage_failures = {
        horizon: value
        for horizon, value in coverage_failures.items()
        if value
    }
    passed = not invalid and not entirely_missing and not coverage_failures
    report = {
        "contract": "solar-future-weather-preflight.v2",
        "status": "PASS" if passed else "FAIL",
        "future_weather_contract": FUTURE_WEATHER_CONTRACT,
        "experiment_config": str(config_path),
        "horizons": reports,
        "invalid_features": invalid,
        "entirely_missing_features": entirely_missing,
        "coverage_failures": coverage_failures,
        "minimum_required_coverage": float(minimum_coverage),
        "cross_horizon_key_equivalence_required": False,
        "cross_horizon_note": (
            "Different forecast horizons may use different source models and "
            "usable periods; equality is required within each model comparison, "
            "not between 24h and 72h."
        ),
        "historical_alignment": {
            "temperature": "Celsius_to_Celsius",
            "humidity": "percent_to_percent",
            "precipitation": "mm_to_mm",
            "total_cloud_cover": "forecast_percent_divided_by_10_to_ASOS_tenths",
            "wind_speed": "mps_to_mps",
            "solar_irradiance": (
                "MSM-only forecast hourly mean W_m2 times 0.0036 to MJ_m2"
            ),
            "sunshine": "MSM-only forecast seconds divided by 3600 to hours",
            "dni_dhi": "MSM ablation only; no direct current ASOS counterpart",
        },
    }
    return report


def main() -> None:
    args = _parser().parse_args()
    config_path = Path(args.config)
    if not config_path.is_absolute():
        config_path = PROJECT_ROOT / config_path
    root = Path(args.root)
    output = Path(args.output)
    if not 0 < args.minimum_coverage <= 1:
        raise ValueError("--minimum-coverage must be in (0, 1]")
    report = validate(
        config_path,
        root=root,
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
