"""ASOS 관측값 이름·타입 표준화와 관측소 좌표 결합."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd

from solar_forecast.collectors.generation_normalizers import read_csv_with_fallback
from solar_forecast.quality.generation_quality import WEATHER_RANGES


WEATHER_COLUMN_MAP = {
    "지점": "station_id",
    "지점명": "station_name",
    "일시": "timestamp",
    "기온(°C)": "temperature_c",
    "강수량(mm)": "precipitation_mm",
    "풍속(m/s)": "wind_speed_mps",
    "습도(%)": "humidity_pct",
    "일조(hr)": "sunshine_hours",
    "일사(MJ/m2)": "solar_irradiance_mj_m2",
    "전운량(10분위)": "total_cloud_cover_tenths",
    "중하층운량(10분위)": "low_mid_cloud_cover_tenths",
}

WEATHER_QC_COLUMN_MAP = {
    "temperature_c": ("taQcflg", "기온 QC플래그"),
    "precipitation_mm": ("rnQcflg", "강수량 QC플래그"),
    "wind_speed_mps": ("wsQcflg", "풍속 QC플래그"),
    "humidity_pct": ("hmQcflg", "습도 QC플래그"),
    "sunshine_hours": ("ssQcflg", "일조 QC플래그"),
    "solar_irradiance_mj_m2": ("icsrQcflg", "일사 QC플래그"),
    "total_cloud_cover_tenths": ("dc10TcaQcflg", "전운량 QC플래그"),
    "low_mid_cloud_cover_tenths": ("dc10LmcsCaQcflg", "중하층운량 QC플래그"),
}

WEATHER_PROVENANCE_COLUMNS = [
    "weather_station_hour_present",
    "station_metadata_status",
    "station_metadata_valid_from",
    "station_metadata_valid_to",
    *[
        f"{column}_{suffix}"
        for column in WEATHER_QC_COLUMN_MAP
        for suffix in ("observed", "invalid", "missing_reason", "qc_flag")
    ],
]

WEATHER_COLUMNS = list(WEATHER_COLUMN_MAP.values()) + [
    "station_latitude",
    "station_longitude",
    "station_elevation_m",
    *WEATHER_PROVENANCE_COLUMNS,
]


class KmaAsosNormalizer:
    """Normalize station-hours before plant joins; retain missingness and QC reasons."""

    def __init__(self, station_metadata_path: Path):
        self.station_metadata_path = Path(station_metadata_path)

    def read(
        self,
        paths: Iterable[Path],
        *,
        station_ids: Iterable[int] | None = None,
    ) -> pd.DataFrame:
        selected = (
            {int(station_id) for station_id in station_ids}
            if station_ids is not None
            else None
        )
        frames: list[pd.DataFrame] = []
        for path in paths:
            frame = read_csv_with_fallback(Path(path))
            frame.columns = [str(column).strip() for column in frame.columns]
            if selected is not None:
                if "지점" not in frame:
                    raise ValueError(f"KMA station column is missing: {path}")
                station = pd.to_numeric(frame["지점"], errors="coerce")
                frame = frame.loc[station.isin(selected)]
            if not frame.empty:
                frames.append(frame)
        if not frames:
            raise ValueError("At least one KMA hourly file is required")
        return self.transform(pd.concat(frames, ignore_index=True))

    def transform(self, frame: pd.DataFrame) -> pd.DataFrame:
        source = frame.copy()
        source.columns = [str(column).strip() for column in source.columns]
        missing = set(WEATHER_COLUMN_MAP) - set(source.columns)
        if missing:
            raise ValueError(f"KMA ASOS columns are missing: {sorted(missing)}")
        result = source[list(WEATHER_COLUMN_MAP)].rename(columns=WEATHER_COLUMN_MAP)
        result["timestamp"] = pd.to_datetime(result["timestamp"], errors="coerce")
        station = pd.to_numeric(result["station_id"], errors="coerce")
        result["station_id"] = station.where(np.isfinite(station) & station.mod(1).eq(0))
        for column, aliases in WEATHER_QC_COLUMN_MAP.items():
            self._normalize_observation(result, source, column, aliases)
        result = result.dropna(subset=["station_id", "timestamp"])
        result["station_id"] = result["station_id"].astype(int)
        result = result.drop_duplicates(["station_id", "timestamp"], keep="last")
        result["weather_station_hour_present"] = True
        result = self._join_station_history(result.reset_index(drop=True))
        return result[WEATHER_COLUMNS].sort_values(["timestamp", "station_id"]).reset_index(drop=True)

    @staticmethod
    def _normalize_observation(
        result: pd.DataFrame, source: pd.DataFrame, column: str, aliases: tuple[str, ...]
    ) -> None:
        raw = result[column]
        present = raw.notna() & raw.astype("string").str.strip().ne("").fillna(False)
        values = pd.to_numeric(raw, errors="coerce")
        reason = pd.Series("observed", index=result.index, dtype="string")
        reason.loc[~present] = "source_missing"
        reason.loc[present & values.isna()] = "non_numeric"
        non_finite = values.notna() & ~np.isfinite(values)
        reason.loc[non_finite] = "non_finite"
        lower, upper = WEATHER_RANGES[column]
        outside = pd.Series(False, index=result.index)
        if lower is not None:
            outside |= values.lt(lower).fillna(False)
        if upper is not None:
            outside |= values.gt(upper).fillna(False)
        reason.loc[outside & ~non_finite] = "out_of_range"

        # Empty QC is unknown, not an error. API names take precedence over
        # equivalent download-column names; retain the provider token verbatim.
        qc = pd.Series(pd.NA, index=result.index, dtype="string")
        for alias in aliases:
            if alias in source:
                incoming = source[alias].astype("string").str.strip().replace("", pd.NA)
                qc = qc.fillna(incoming)
        qc_number = pd.to_numeric(qc, errors="coerce")
        reason.loc[qc_number.eq(1).fillna(False)] = "qc_error"
        reason.loc[qc_number.eq(9).fillna(False)] = "qc_missing"
        observed = reason.eq("observed")
        result[column] = values.where(observed)
        result[f"{column}_observed"] = observed.astype(bool)
        result[f"{column}_invalid"] = reason.isin(
            ["non_numeric", "non_finite", "out_of_range", "qc_error", "qc_missing"]
        )
        result[f"{column}_missing_reason"] = reason
        result[f"{column}_qc_flag"] = qc

    def _join_station_history(self, result: pd.DataFrame) -> pd.DataFrame:
        """Use date-valid metadata; never extrapolate the newest dated record.

        Metadata dates have day precision. End dates are inclusive; where a
        move shares an old end date and a new start date, the newest start wins.
        A single undated record is retained for legacy compatibility and marked
        validity_unknown. Unknown or expired histories never borrow later data.
        """
        metadata = read_csv_with_fallback(self.station_metadata_path)
        metadata.columns = [str(column).strip() for column in metadata.columns]
        required = {"지점", "위도", "경도", "노장해발고도(m)"}
        missing = required - set(metadata.columns)
        if missing:
            raise ValueError(f"KMA station metadata columns are missing: {sorted(missing)}")
        metadata = metadata.rename(columns={
            "지점": "station_id", "위도": "station_latitude",
            "경도": "station_longitude", "노장해발고도(m)": "station_elevation_m",
        })
        coordinates = ["station_latitude", "station_longitude", "station_elevation_m"]
        for column in ["station_id", *coordinates]:
            values = pd.to_numeric(metadata[column], errors="coerce")
            metadata[column] = values.where(np.isfinite(values))
        metadata = metadata.dropna(subset=["station_id"])
        metadata = metadata.loc[metadata["station_id"].mod(1).eq(0)].copy()
        metadata["station_id"] = metadata["station_id"].astype(int)
        for column, bound in (("station_latitude", 90), ("station_longitude", 180)):
            metadata[column] = metadata[column].where(metadata[column].abs().le(bound))
        for source_column, target in (
            ("시작일", "station_metadata_valid_from"),
            ("종료일", "station_metadata_valid_to"),
        ):
            raw = metadata.get(source_column, pd.Series(pd.NA, index=metadata.index))
            metadata[target] = pd.to_datetime(raw, errors="coerce").dt.normalize()
            metadata[f"_{target}_malformed"] = (
                raw.notna() & raw.astype("string").str.strip().ne("").fillna(False)
                & metadata[target].isna()
            )
        result[coordinates] = np.nan
        result["station_metadata_status"] = "station_not_found"
        result["station_metadata_valid_from"] = pd.NaT
        result["station_metadata_valid_to"] = pd.NaT
        for station_id, rows in result.groupby("station_id", sort=False):
            history = metadata.loc[metadata["station_id"].eq(station_id)].copy()
            if history.empty:
                continue
            start = "station_metadata_valid_from"
            end = "station_metadata_valid_to"
            malformed = history[f"_{start}_malformed"] | history[f"_{end}_malformed"]
            malformed |= history[end].lt(history[start]).fillna(False)
            if len(history) == 1 and history[start].isna().all() and not malformed.any():
                record = history.iloc[0]
                if pd.isna(record[end]):
                    result.loc[rows.index, coordinates] = record[coordinates].to_numpy()
                    result.loc[rows.index, "station_metadata_status"] = "validity_unknown"
                    continue
            history = history.loc[~malformed & history[start].notna()]
            if history.empty:
                result.loc[rows.index, "station_metadata_status"] = "validity_unknown"
                continue
            # Overlapping historical records are resolved by the most recent
            # start, consistently for every plant sharing this station-hour.
            history = history.sort_values(start, kind="stable").drop_duplicates(start, keep="last")
            selected = pd.merge_asof(
                rows[["timestamp"]].assign(_row=rows.index).sort_values("timestamp"),
                history[[start, end, *coordinates]],
                left_on="timestamp", right_on=start, direction="backward",
            ).set_index("_row")
            covered = selected[start].notna() & (
                selected[end].isna() | selected["timestamp"].dt.normalize().le(selected[end])
            )
            result.loc[rows.index, "station_metadata_status"] = "outside_validity"
            indices = selected.index[covered]
            result.loc[indices, [*coordinates, start, end]] = selected.loc[indices, [*coordinates, start, end]]
            result.loc[indices, "station_metadata_status"] = "matched"
        return result
