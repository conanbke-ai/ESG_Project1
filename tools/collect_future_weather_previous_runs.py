"""Collect leakage-safe fixed-lead weather forecasts for 24h/72h benchmarks.

Source: Open-Meteo Previous Runs API using JMA MSM for Korea.
The API's *_previous_day1 and *_previous_day3 fields represent forecasts made
24 and 72 hours before valid time. Raw responses are cached per plant/year/lead
so interrupted collection can resume without repeating successful downloads.
"""
from __future__ import annotations

import argparse
from datetime import date
import json
from pathlib import Path
import time

import pandas as pd
import requests

from solar_forecast.features.future_weather import (
    FUTURE_WEATHER_FEATURES,
    FUTURE_WEATHER_CONTRACT,
)


ENDPOINT = "https://previous-runs-api.open-meteo.com/v1/forecast"
MODEL = "jma_msm"
LEAD_SUFFIX = {24: "previous_day1", 72: "previous_day3"}
SOURCE_VARIABLES = {
    "future_temperature_c": "temperature_2m",
    "future_humidity_pct": "relative_humidity_2m",
    "future_precipitation_mm": "precipitation",
    "future_cloud_cover_pct": "cloud_cover",
    "future_wind_speed_mps": "wind_speed_10m",
    "future_ghi_w_m2": "shortwave_radiation",
    "future_dni_w_m2": "direct_normal_irradiance",
    "future_dhi_w_m2": "diffuse_radiation",
}


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--registry",
        default="file/standardized/plant_registry.csv",
        help="Plant registry CSV containing plant_id, latitude and longitude.",
    )
    parser.add_argument(
        "--output-root",
        default="file/forecast_weather",
    )
    parser.add_argument("--start-year", type=int, default=2022)
    parser.add_argument("--end-year", type=int, default=2025)
    parser.add_argument(
        "--horizons",
        nargs="+",
        type=int,
        default=[24, 72],
        choices=[24, 72],
    )
    parser.add_argument("--timeout", type=int, default=120)
    parser.add_argument("--retries", type=int, default=5)
    parser.add_argument("--sleep-seconds", type=float, default=0.5)
    return parser


def _registry(path: Path) -> pd.DataFrame:
    frame = pd.read_csv(path, dtype={"plant_id": str}, low_memory=False)
    required = {"plant_id", "latitude", "longitude"}
    missing = required - set(frame)
    if missing:
        raise ValueError(f"Plant registry columns are missing: {sorted(missing)}")
    frame["latitude"] = pd.to_numeric(frame["latitude"], errors="coerce")
    frame["longitude"] = pd.to_numeric(frame["longitude"], errors="coerce")
    if "model_ready_status" in frame:
        frame = frame.loc[frame["model_ready_status"].eq("eligible")]
    frame = frame.dropna(subset=["plant_id", "latitude", "longitude"]).copy()
    frame = frame.loc[
        frame["latitude"].between(32.0, 39.5)
        & frame["longitude"].between(124.0, 132.5)
    ]
    eligible_ids = set(
        frame.loc[
            frame.get(
                "model_ready_status",
                pd.Series("eligible", index=frame.index),
            ).eq("eligible"),
            "plant_id",
        ].astype(str)
    )
    usable = frame.dropna(subset=["plant_id", "latitude", "longitude"]).copy()
    usable = usable.loc[
        usable["latitude"].between(32.0, 39.5)
        & usable["longitude"].between(124.0, 132.5)
    ]
    usable = usable.drop_duplicates("plant_id", keep="last")
    missing_ids = sorted(eligible_ids - set(usable["plant_id"].astype(str)))
    if missing_ids:
        raise ValueError(
            "Eligible plants are missing plant coordinates; future weather "
            f"will not fall back to region or ASOS station: {missing_ids}"
        )
    if usable.empty:
        raise ValueError("No eligible plants have usable coordinates")
    usable["coordinate_source"] = "plant_registry_coordinates"
    return usable.sort_values("plant_id", kind="stable").reset_index(drop=True)


def _request_json(
    session: requests.Session,
    *,
    latitude: float,
    longitude: float,
    start_date: str,
    end_date: str,
    horizon: int,
    timeout: int,
    retries: int,
) -> dict:
    suffix = LEAD_SUFFIX[horizon]
    hourly = [f"{name}_{suffix}" for name in SOURCE_VARIABLES.values()]
    params = {
        "latitude": latitude,
        "longitude": longitude,
        "start_date": start_date,
        "end_date": end_date,
        "hourly": ",".join(hourly),
        "models": MODEL,
        "timezone": "Asia/Seoul",
        "wind_speed_unit": "ms",
        "temporal_resolution": "hourly",
        "cell_selection": "nearest",
    }
    last_error: Exception | None = None
    for attempt in range(retries):
        try:
            response = session.get(
                ENDPOINT,
                params=params,
                timeout=timeout,
                headers={"User-Agent": "ESG_Project1 solar benchmark collector"},
            )
            if response.status_code == 429:
                delay = min(60.0, 2.0 ** (attempt + 1))
                time.sleep(delay)
                continue
            response.raise_for_status()
            payload = response.json()
            if payload.get("error"):
                raise RuntimeError(str(payload))
            return payload
        except (requests.RequestException, ValueError, RuntimeError) as exc:
            last_error = exc
            if attempt + 1 < retries:
                time.sleep(min(30.0, 2.0 ** attempt))
    raise RuntimeError(
        f"Previous Runs request failed after {retries} attempts"
    ) from last_error


def _normalize(
    payload: dict,
    *,
    plant_id: str,
    plant_latitude: float,
    plant_longitude: float,
    horizon: int,
) -> pd.DataFrame:
    hourly = payload.get("hourly")
    if not isinstance(hourly, dict) or "time" not in hourly:
        raise ValueError("Previous Runs response omitted hourly data")
    suffix = LEAD_SUFFIX[horizon]
    result = pd.DataFrame(
        {
            "timestamp": pd.to_datetime(hourly["time"], errors="raise"),
            "plant_id": str(plant_id),
        }
    )
    for output_name, source_name in SOURCE_VARIABLES.items():
        key = f"{source_name}_{suffix}"
        values = hourly.get(key)
        if values is None:
            result[output_name] = pd.NA
        else:
            result[output_name] = pd.to_numeric(
                pd.Series(values),
                errors="coerce",
            )
    result["forecast_origin"] = (
        result["timestamp"] - pd.Timedelta(hours=horizon)
    )
    result["plant_latitude"] = float(plant_latitude)
    result["plant_longitude"] = float(plant_longitude)
    result["grid_latitude"] = float(payload.get("latitude", plant_latitude))
    result["grid_longitude"] = float(payload.get("longitude", plant_longitude))
    result["coordinate_source"] = "plant_registry_coordinates"
    result["cell_selection"] = "nearest"
    result["horizon_hours"] = int(horizon)
    result["forecast_model"] = MODEL
    result["forecast_source"] = "open_meteo_previous_runs"
    return result


def _year_bounds(year: int) -> tuple[str, str]:
    return date(year, 1, 1).isoformat(), date(year, 12, 31).isoformat()


def collect(args: argparse.Namespace) -> None:
    registry_path = Path(args.registry)
    output_root = Path(args.output_root)
    output_root.mkdir(parents=True, exist_ok=True)
    cache = output_root / "raw"
    cache.mkdir(parents=True, exist_ok=True)
    plants = _registry(registry_path)
    session = requests.Session()

    for horizon in args.horizons:
        parts: list[pd.DataFrame] = []
        for row in plants.itertuples(index=False):
            for year in range(args.start_year, args.end_year + 1):
                cached = cache / f"{row.plant_id.replace(':', '__')}_{year}_{horizon}h.csv.gz"
                if cached.is_file():
                    part = pd.read_csv(
                        cached,
                        dtype={"plant_id": str},
                        parse_dates=["timestamp", "forecast_origin"],
                    )
                else:
                    start_date, end_date = _year_bounds(year)
                    payload = _request_json(
                        session,
                        latitude=float(row.latitude),
                        longitude=float(row.longitude),
                        start_date=start_date,
                        end_date=end_date,
                        horizon=horizon,
                        timeout=args.timeout,
                        retries=args.retries,
                    )
                    part = _normalize(
                        payload,
                        plant_id=str(row.plant_id),
                        plant_latitude=float(row.latitude),
                        plant_longitude=float(row.longitude),
                        horizon=horizon,
                    )
                    part.to_csv(
                        cached,
                        index=False,
                        compression={"method": "gzip", "compresslevel": 1, "mtime": 1},
                    )
                    time.sleep(args.sleep_seconds)
                parts.append(part)

        combined = pd.concat(parts, ignore_index=True)
        combined = combined.drop_duplicates(
            ["plant_id", "timestamp", "forecast_origin", "horizon_hours"],
            keep="last",
        ).sort_values(["timestamp", "plant_id"], kind="stable")
        output = output_root / f"open_meteo_jma_msm_{horizon}h.csv.gz"
        combined.to_csv(
            output,
            index=False,
            compression={"method": "gzip", "compresslevel": 1, "mtime": 1},
        )
        manifest = {
            "contract": FUTURE_WEATHER_CONTRACT,
            "source": "Open-Meteo Previous Runs API",
            "endpoint": ENDPOINT,
            "model": MODEL,
            "horizon_hours": horizon,
            "lead_field_suffix": LEAD_SUFFIX[horizon],
            "rows": int(len(combined)),
            "plants": int(combined["plant_id"].nunique()),
            "start": combined["timestamp"].min().isoformat(),
            "end": combined["timestamp"].max().isoformat(),
            "features": list(FUTURE_WEATHER_FEATURES),
            "forecast_origin_rule": "target_timestamp_minus_fixed_lead",
            "spatial_contract": "plant_registry_coordinates_only",
            "coordinate_source": "plant_registry_coordinates",
            "cell_selection": "nearest",
            "region_used_for_weather_lookup": False,
            "asos_station_used_for_weather_lookup": False,
            "output": str(output),
        }
        (output_root / f"open_meteo_jma_msm_{horizon}h.manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(
            f"{horizon}h future weather: {len(combined):,} rows -> {output}",
            flush=True,
        )


if __name__ == "__main__":
    collect(_parser().parse_args())
