from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
import warnings
from typing import Iterable

import numpy as np
import pandas as pd


WEATHER_RANGES: dict[str, tuple[float | None, float | None]] = {
    "temperature_c": (-60.0, 60.0),
    "precipitation_mm": (0.0, None),
    "sunshine_hours": (0.0, 1.0),
    "solar_irradiance_mj_m2": (0.0, None),
    "wind_speed_mps": (0.0, None),
    "humidity_pct": (0.0, 100.0),
    "total_cloud_cover_tenths": (0.0, 10.0),
    "low_mid_cloud_cover_tenths": (0.0, 10.0),
}

QUALITY_COLUMNS = [
    "capacity_factor",
    "quality_code",
    "quality_train_eligible",
    "quality_review_required",
    "quality_missing_generation",
    "quality_negative_generation",
    "quality_capacity_exceeded",
    "quality_daylight_zero",
    "quality_flatline",
    "quality_daily_aggregate_profile",
    "quality_invalid_weather",
    "quality_missing_weather",
]


@dataclass(frozen=True)
class PhysicalQualityConfig:
    """Conservative rules that separate impossible values from suspicious context."""

    capacity_tolerance: float = 1.2
    daylight_irradiance_threshold: float = 0.05
    flatline_min_hours: int = 4
    aggregate_profile_min_active_days: int = 30
    aggregate_profile_single_positive_ratio: float = 0.95
    aggregate_profile_dominant_hour_ratio: float = 0.95
    aggregate_profile_night_hours: tuple[int, ...] = (22, 23, 0, 1, 2, 3, 4, 5)

    def __post_init__(self) -> None:
        if self.capacity_tolerance <= 1:
            raise ValueError("capacity_tolerance must be greater than 1")
        if self.daylight_irradiance_threshold < 0:
            raise ValueError("daylight_irradiance_threshold cannot be negative")
        if self.flatline_min_hours < 2:
            raise ValueError("flatline_min_hours must be at least two")
        if self.aggregate_profile_min_active_days < 1:
            raise ValueError("aggregate_profile_min_active_days must be positive")
        for name, value in (
            (
                "aggregate_profile_single_positive_ratio",
                self.aggregate_profile_single_positive_ratio,
            ),
            (
                "aggregate_profile_dominant_hour_ratio",
                self.aggregate_profile_dominant_hour_ratio,
            ),
        ):
            if not 0 < value <= 1:
                raise ValueError(f"{name} must be in (0, 1]")
        if (
            not self.aggregate_profile_night_hours
            or len(set(self.aggregate_profile_night_hours))
            != len(self.aggregate_profile_night_hours)
            or any(hour < 0 or hour > 23 for hour in self.aggregate_profile_night_hours)
        ):
            raise ValueError(
                "aggregate_profile_night_hours must contain unique hours between 0 and 23"
            )


class GenerationQualityPolicy:
    """Add auditable physical/contextual flags without fabricating target values."""

    def __init__(self, config: PhysicalQualityConfig | None = None):
        self.config = config or PhysicalQualityConfig()

    def apply(self, frame: pd.DataFrame) -> pd.DataFrame:
        required = {"timestamp", "plant_id", "generation_mwh"}
        missing = required - set(frame.columns)
        if missing:
            raise ValueError(f"Quality policy columns are missing: {sorted(missing)}")

        result = frame.copy()
        result["timestamp"] = pd.to_datetime(result["timestamp"], errors="coerce")
        result["generation_mwh"] = pd.to_numeric(result["generation_mwh"], errors="coerce")
        if result["timestamp"].isna().any():
            raise ValueError("Quality policy input contains invalid timestamps")
        if "energy_source" not in result:
            result["energy_source"] = "unknown"

        result["quality_missing_generation"] = result["generation_mwh"].isna()
        result["quality_negative_generation"] = result["generation_mwh"].lt(0).fillna(False)

        if "capacity_mw" in result:
            capacity = pd.to_numeric(result["capacity_mw"], errors="coerce")
        else:
            capacity = pd.Series(np.nan, index=result.index, dtype=float)
        valid_capacity = capacity.gt(0)
        result["capacity_factor"] = result["generation_mwh"].div(capacity).where(valid_capacity)
        result["quality_capacity_exceeded"] = (
            result["capacity_factor"].gt(self.config.capacity_tolerance).fillna(False)
        )

        weather_flags: list[pd.Series] = []
        weather_missing: list[pd.Series] = []
        for column, (lower, upper) in WEATHER_RANGES.items():
            if column not in result:
                continue
            values = pd.to_numeric(result[column], errors="coerce")
            invalid = pd.Series(False, index=result.index)
            if lower is not None:
                invalid |= values.lt(lower).fillna(False)
            if upper is not None:
                invalid |= values.gt(upper).fillna(False)
            weather_flags.append(invalid)
            weather_missing.append(values.isna())
            # Invalid observations are missing, never zero. The flag retains
            # the reason so a later train-fitted imputer can handle them.
            result[column] = values.mask(invalid)
        result["quality_invalid_weather"] = (
            pd.concat(weather_flags, axis=1).any(axis=1)
            if weather_flags
            else False
        )
        result["quality_missing_weather"] = (
            pd.concat(weather_missing, axis=1).any(axis=1)
            if weather_missing
            else False
        )

        irradiance = (
            pd.to_numeric(result["solar_irradiance_mj_m2"], errors="coerce")
            if "solar_irradiance_mj_m2" in result
            else pd.Series(np.nan, index=result.index, dtype=float)
        )
        daylight = irradiance.ge(self.config.daylight_irradiance_threshold).fillna(False)
        solar = result["energy_source"].astype(str).eq("solar")
        result["quality_daylight_zero"] = (
            solar & daylight & result["generation_mwh"].eq(0).fillna(False)
        )
        result["quality_flatline"] = self._flatline_flags(result, daylight)
        result["quality_daily_aggregate_profile"] = self._daily_aggregate_profile_flags(
            result
        )

        # Impossible target values and source-level daily aggregates cannot be
        # valid hourly labels. Capacity excess, daylight zero and flatline stay
        # review-only because metadata errors, curtailment and planned outages
        # can look identical at row level.
        result["quality_train_eligible"] = ~(
            result["quality_missing_generation"]
            | result["quality_negative_generation"]
            | result["quality_daily_aggregate_profile"]
        )
        review_columns = [
            "quality_capacity_exceeded",
            "quality_daylight_zero",
            "quality_flatline",
            "quality_daily_aggregate_profile",
            "quality_invalid_weather",
            "quality_missing_weather",
        ]
        result["quality_review_required"] = result[review_columns].any(axis=1)
        result["quality_code"] = self._quality_codes(result)
        return result

    def _flatline_flags(self, frame: pd.DataFrame, daylight: pd.Series) -> pd.Series:
        flags = pd.Series(False, index=frame.index)
        working = frame[["plant_id", "timestamp", "generation_mwh"]].copy()
        working["_daylight"] = daylight
        working["_original_index"] = frame.index
        working = working.sort_values(["plant_id", "timestamp"], kind="stable")
        for _, group in working.groupby("plant_id", sort=False, observed=True):
            consecutive = group["timestamp"].diff().eq(pd.Timedelta(hours=1))
            same_value = group["generation_mwh"].eq(group["generation_mwh"].shift())
            eligible = group["_daylight"] & group["generation_mwh"].gt(0)
            new_run = ~(consecutive & same_value & eligible & eligible.shift(fill_value=False))
            run_id = new_run.cumsum()
            run_size = group.groupby(run_id, sort=False)["generation_mwh"].transform("size")
            group_flags = eligible & run_size.ge(self.config.flatline_min_hours)
            flags.loc[group.loc[group_flags, "_original_index"]] = True
        return flags

    def _daily_aggregate_profile_flags(self, frame: pd.DataFrame) -> pd.Series:
        """Flag solar plants whose daily totals were placed in one night bucket.

        A few public exports expose a daily total through an hourly-shaped file.
        Reconstructing the missing intraday curve would fabricate target labels,
        so every row for a decisively detected plant is excluded. The combined
        thresholds deliberately avoid treating isolated short winter days,
        outages, or sparse sensors as daily aggregates.
        """

        flags = pd.Series(False, index=frame.index)
        solar = frame["energy_source"].astype(str).eq("solar")
        profile = self._daily_aggregate_profile_evidence(frame)
        suspicious = profile.index[profile["detected"]]
        if len(suspicious):
            flags.loc[solar & frame["plant_id"].isin(suspicious)] = True
        return flags

    def _daily_aggregate_profile_evidence(self, frame: pd.DataFrame) -> pd.DataFrame:
        columns = [
            "aggregate_active_days",
            "aggregate_single_positive_day_ratio",
            "aggregate_dominant_hour",
            "aggregate_dominant_hour_ratio",
            "detected",
        ]
        solar = frame["energy_source"].astype(str).eq("solar")
        working = frame.loc[solar, ["plant_id", "timestamp", "generation_mwh"]].copy()
        if working.empty:
            return pd.DataFrame(columns=columns)
        working["_date"] = working["timestamp"].dt.normalize()
        working["_hour"] = working["timestamp"].dt.hour
        working["_generation"] = pd.to_numeric(
            working["generation_mwh"], errors="coerce"
        )
        positive_hours = working.loc[
            working["_generation"].gt(0), ["plant_id", "_date", "_hour"]
        ].drop_duplicates(["plant_id", "_date", "_hour"])
        if positive_hours.empty:
            return pd.DataFrame(columns=columns)

        daily = (
            positive_hours.groupby(["plant_id", "_date"], observed=True)
            .size()
            .rename("positive_hours")
            .reset_index()
        )
        active_days = daily.groupby("plant_id", observed=True).size().rename("active_days")
        single_days = daily.loc[daily["positive_hours"].eq(1), ["plant_id", "_date"]]
        if single_days.empty:
            profile = active_days.to_frame()
            profile["single_days"] = 0
            profile["dominant_days"] = 0
            profile["_hour"] = np.nan
        else:
            single_counts = (
                single_days.groupby("plant_id", observed=True).size().rename("single_days")
            )
            single_hours = single_days.merge(
                positive_hours,
                on=["plant_id", "_date"],
                how="inner",
                validate="one_to_many",
            )
            dominant = (
                single_hours.groupby(["plant_id", "_hour"], observed=True)
                .size()
                .rename("dominant_days")
                .reset_index()
                .sort_values(
                    ["plant_id", "dominant_days", "_hour"],
                    ascending=[True, False, True],
                    kind="stable",
                )
                .drop_duplicates("plant_id", keep="first")
                .set_index("plant_id")
            )
            profile = pd.concat([active_days, single_counts], axis=1).fillna(0)
            profile = profile.join(dominant[["dominant_days", "_hour"]], how="left")
        profile["single_positive_ratio"] = profile["single_days"].div(
            profile["active_days"]
        )
        # Divide by all active days, not only single-positive days. This makes
        # the dominant-hour threshold an independent, stricter safeguard.
        profile["dominant_hour_ratio"] = profile["dominant_days"].div(
            profile["active_days"]
        )
        profile["detected"] = (
            profile["active_days"].ge(self.config.aggregate_profile_min_active_days)
            & profile["single_positive_ratio"].ge(
                self.config.aggregate_profile_single_positive_ratio
            )
            & profile["dominant_hour_ratio"].ge(
                self.config.aggregate_profile_dominant_hour_ratio
            )
            & profile["_hour"].isin(self.config.aggregate_profile_night_hours)
        )
        return profile.rename(
            columns={
                "active_days": "aggregate_active_days",
                "single_positive_ratio": "aggregate_single_positive_day_ratio",
                "_hour": "aggregate_dominant_hour",
                "dominant_hour_ratio": "aggregate_dominant_hour_ratio",
            }
        )[columns]

    @staticmethod
    def _quality_codes(frame: pd.DataFrame) -> pd.Series:
        pairs = (
            ("quality_missing_generation", "missing_generation"),
            ("quality_negative_generation", "negative_generation"),
            ("quality_capacity_exceeded", "capacity_exceeded"),
            ("quality_daylight_zero", "daylight_zero"),
            ("quality_flatline", "flatline"),
            ("quality_daily_aggregate_profile", "daily_aggregate_profile"),
            ("quality_invalid_weather", "invalid_weather"),
            ("quality_missing_weather", "missing_weather"),
        )
        codes = np.full(len(frame), "", dtype=object)
        for column, label in pairs:
            mask = frame[column].to_numpy(dtype=bool)
            codes[mask] = np.where(codes[mask] == "", label, codes[mask] + "|" + label)
        codes[codes == ""] = "ok"
        return pd.Series(codes, index=frame.index, dtype="string")


class PlantQualityProfiler:
    """Summarize sensor risk per plant without declaring equipment failure."""

    def __init__(self, policy: GenerationQualityPolicy | None = None):
        self.policy = policy or GenerationQualityPolicy()

    def profile(self, frame: pd.DataFrame) -> pd.DataFrame:
        audited = self.policy.apply(frame)
        temporal_consistency = self._temporal_profile_consistency(audited)
        peer_consistency = self._peer_pattern_consistency(audited)
        aggregate_evidence = self.policy._daily_aggregate_profile_evidence(audited)
        rows: list[dict[str, object]] = []
        for plant_id, group in audited.groupby("plant_id", sort=True, observed=True):
            start, end = group["timestamp"].min(), group["timestamp"].max()
            expected = max(1, int((end - start) / pd.Timedelta(hours=1)) + 1)
            aggregate = (
                aggregate_evidence.loc[plant_id]
                if plant_id in aggregate_evidence.index
                else None
            )
            row: dict[str, object] = {
                "company": str(group["company"].iloc[0]) if "company" in group else "unknown",
                "plant_id": str(plant_id),
                "plant": str(group["plant"].iloc[0]) if "plant" in group else str(plant_id),
                "region": str(group["region"].iloc[0]) if "region" in group else "unknown",
                "energy_source": str(group["energy_source"].mode().iloc[0]),
                "rows": len(group),
                "start": start.isoformat(),
                "end": end.isoformat(),
                "hourly_coverage": min(1.0, len(group) / expected),
                "capacity_coverage": float(group["capacity_factor"].notna().mean()),
                "missing_generation_rate": float(group["quality_missing_generation"].mean()),
                "negative_generation_rate": float(group["quality_negative_generation"].mean()),
                "capacity_exceeded_rate": float(group["quality_capacity_exceeded"].mean()),
                "daylight_zero_rate": self._conditional_rate(
                    group["quality_daylight_zero"],
                    pd.to_numeric(group.get("solar_irradiance_mj_m2"), errors="coerce").ge(
                        self.policy.config.daylight_irradiance_threshold
                    ) if "solar_irradiance_mj_m2" in group else pd.Series(False, index=group.index),
                ),
                "flatline_rate": float(group["quality_flatline"].mean()),
                "positive_flatline_rate": self._positive_flatline_rate(group),
                "daily_aggregate_profile_detected": bool(
                    group["quality_daily_aggregate_profile"].any()
                ),
                "aggregate_active_days": (
                    int(aggregate["aggregate_active_days"])
                    if aggregate is not None
                    else 0
                ),
                "aggregate_single_positive_day_ratio": (
                    float(aggregate["aggregate_single_positive_day_ratio"])
                    if aggregate is not None
                    else np.nan
                ),
                "aggregate_dominant_hour": (
                    int(aggregate["aggregate_dominant_hour"])
                    if aggregate is not None
                    and pd.notna(aggregate["aggregate_dominant_hour"])
                    else np.nan
                ),
                "aggregate_dominant_hour_ratio": (
                    float(aggregate["aggregate_dominant_hour_ratio"])
                    if aggregate is not None
                    else np.nan
                ),
                "invalid_weather_rate": float(group["quality_invalid_weather"].mean()),
                "missing_weather_rate": float(group["quality_missing_weather"].mean()),
                "temporal_profile_consistency": temporal_consistency.get(str(plant_id), np.nan),
                "peer_pattern_correlation": peer_consistency.get(str(plant_id), np.nan),
            }
            row["sensor_risk"] = self._sensor_risk(row)
            row["recommended_action"] = self._recommendation(row)
            rows.append(row)
        return pd.DataFrame(rows)

    def _positive_flatline_rate(self, group: pd.DataFrame) -> float:
        working = group.copy()
        if "plant_id" not in working:
            working["plant_id"] = working.get("plant", "reference")
        daylight = pd.Series(True, index=working.index)
        return float(self.policy._flatline_flags(working, daylight).mean())

    @staticmethod
    def _conditional_rate(flag: pd.Series, condition: pd.Series) -> float:
        condition = condition.fillna(False)
        return float(flag[condition].mean()) if condition.any() else float("nan")

    @staticmethod
    def _temporal_profile_consistency(frame: pd.DataFrame) -> dict[str, float]:
        scores: dict[str, float] = {}
        working = frame.assign(date=frame["timestamp"].dt.date, hour=frame["timestamp"].dt.hour)
        for plant_id, group in working.groupby("plant_id", sort=False, observed=True):
            pivot = group.pivot_table(
                index="date",
                columns="hour",
                values="generation_mwh",
                aggfunc="mean",
                observed=True,
            )
            if len(pivot) < 7:
                scores[str(plant_id)] = float("nan")
                continue
            daily_max = pivot.max(axis=1).replace(0, np.nan)
            normalized = pivot.div(daily_max, axis=0)
            reference = normalized.median(axis=0)
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", RuntimeWarning)
                correlations = normalized.apply(lambda row: row.corr(reference), axis=1).dropna()
            scores[str(plant_id)] = float(correlations.median()) if not correlations.empty else float("nan")
        return scores

    @staticmethod
    def _peer_pattern_consistency(frame: pd.DataFrame) -> dict[str, float]:
        scores: dict[str, float] = {}
        if "region" not in frame:
            return scores
        peer_group_columns = ["region", "energy_source"] if "energy_source" in frame else ["region"]
        for _, group in frame.groupby(peer_group_columns, sort=False, observed=True):
            if group["plant_id"].nunique() < 2:
                continue
            values = group["capacity_factor"].copy()
            fallback_scale = group.groupby("plant_id", observed=True)["generation_mwh"].transform(
                lambda series: series.quantile(0.95)
            ).replace(0, np.nan)
            values = values.where(values.notna(), group["generation_mwh"].div(fallback_scale))
            normalized = group.assign(_normalized_generation=values)
            pivot = normalized.pivot_table(
                index="timestamp",
                columns="plant_id",
                values="_normalized_generation",
                aggfunc="mean",
                observed=True,
            )
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", RuntimeWarning)
                correlations = pivot.corr(min_periods=168)
            for plant_id in correlations.columns:
                peers = correlations.loc[plant_id].drop(labels=[plant_id], errors="ignore").dropna()
                scores[str(plant_id)] = float(peers.median()) if not peers.empty else float("nan")
        return scores

    @staticmethod
    def _sensor_risk(row: dict[str, object]) -> str:
        if (
            bool(row["daily_aggregate_profile_detected"])
            or float(row["missing_generation_rate"]) >= 0.10
            or float(row["flatline_rate"]) >= 0.05
            or float(row["negative_generation_rate"]) > 0
            or float(row["hourly_coverage"]) < 0.80
        ):
            return "high"
        temporal = float(row["temporal_profile_consistency"])
        peer = float(row["peer_pattern_correlation"])
        if (
            float(row["missing_generation_rate"]) >= 0.01
            or float(row["flatline_rate"]) >= 0.01
            or float(row["capacity_exceeded_rate"]) >= 0.01
            or float(row["invalid_weather_rate"]) >= 0.01
            or (not np.isnan(temporal) and temporal < 0.40)
            or (not np.isnan(peer) and peer < 0.20)
        ):
            return "review"
        return "low"

    @staticmethod
    def _recommendation(row: dict[str, object]) -> str:
        if bool(row["daily_aggregate_profile_detected"]):
            return "exclude plant from hourly training; retain as daily aggregate evidence"
        if (
            float(row["negative_generation_rate"]) > 0
            or float(row["missing_generation_rate"]) >= 0.10
        ):
            return "exclude flagged target intervals; verify meter/status data before imputation"
        if float(row["flatline_rate"]) >= 0.05:
            return "quarantine flatline intervals until raw meter and peer-series review"
        if float(row["capacity_exceeded_rate"]) >= 0.01:
            return "verify capacity metadata and interval/unit before filtering"
        if float(row["flatline_rate"]) >= 0.01:
            return "compare flatline intervals with peer plants and irradiance"
        if row["sensor_risk"] == "review":
            return "retain with quality mask; review contextual anomalies separately"
        return "retain; monitor quality rates after each collection"


@dataclass(frozen=True)
class QualityAuditResult:
    report_path: Path
    manifest_path: Path
    plants: int
    high_risk_plants: int
    review_plants: int
    preprocessing_artifact_plants: int


class QualityAuditService:
    """Application service that persists plant diagnostics and their thresholds."""

    def __init__(self, profiler: PlantQualityProfiler | None = None):
        self.profiler = profiler or PlantQualityProfiler()

    def run(
        self,
        frame: pd.DataFrame,
        output_dir: Path,
        *,
        reference_paths: Iterable[Path] | None = None,
        artifact_name: str = "plant_quality",
    ) -> QualityAuditResult:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        report = self.profiler.profile(frame)
        if reference_paths:
            reference = self._reference_flatline_evidence(report, reference_paths)
            report = report.merge(reference, on=["company", "plant"], how="left", validate="one_to_one")
            report["preprocessing_artifact_suspected"] = (
                report["positive_flatline_rate"].ge(0.01)
                & report["reference_positive_flatline_rate"].le(0.001)
                & report["reference_rows"].ge(168)
            ).fillna(False)
            artifact = report["preprocessing_artifact_suspected"]
            report.loc[artifact, "sensor_risk"] = "pipeline_artifact"
            report.loc[artifact, "recommended_action"] = (
                "rebuild from standardized official raw; repeated positives are absent in raw data"
            )
        else:
            report["preprocessing_artifact_suspected"] = False
        if not artifact_name.replace("_", "").isalnum():
            raise ValueError("artifact_name must contain only letters, numbers, and underscores")
        report_path = output_dir / f"{artifact_name}_report.csv"
        report.to_csv(report_path, index=False, encoding="utf-8-sig")
        high = int(report["sensor_risk"].eq("high").sum())
        review = int(report["sensor_risk"].eq("review").sum())
        artifacts = int(report["preprocessing_artifact_suspected"].sum())
        aggregate_profile = report["daily_aggregate_profile_detected"].fillna(False).astype(bool)
        manifest = {
            "created_at": datetime.now().isoformat(),
            "report": str(report_path),
            "plants": len(report),
            "sensor_risk": report["sensor_risk"].value_counts().to_dict(),
            "preprocessing_artifact_plants": artifacts,
            "daily_aggregate_profile_plants": {
                "count": int(aggregate_profile.sum()),
                "plant_ids": sorted(
                    report.loc[aggregate_profile, "plant_id"].astype(str).tolist()
                ),
            },
            "policy": {
                "capacity_tolerance": self.profiler.policy.config.capacity_tolerance,
                "daylight_irradiance_threshold": self.profiler.policy.config.daylight_irradiance_threshold,
                "flatline_min_hours": self.profiler.policy.config.flatline_min_hours,
                "daily_aggregate_profile": {
                    "minimum_active_days": self.profiler.policy.config.aggregate_profile_min_active_days,
                    "minimum_single_positive_day_ratio": self.profiler.policy.config.aggregate_profile_single_positive_ratio,
                    "minimum_dominant_night_hour_ratio": self.profiler.policy.config.aggregate_profile_dominant_hour_ratio,
                    "night_hours": list(
                        self.profiler.policy.config.aggregate_profile_night_hours
                    ),
                    "action": "exclude_entire_plant_from_hourly_training",
                },
                "negative_generation": "invalid_missing_not_zero",
                "target_missing": "not_imputed_by_quality_policy",
                "contextual_flags": "not_automatically_removed_except_daily_aggregate_profile",
            },
        }
        manifest_path = output_dir / (
            "quality_manifest.json" if artifact_name == "plant_quality" else f"{artifact_name}_manifest.json"
        )
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        return QualityAuditResult(report_path, manifest_path, len(report), high, review, artifacts)

    def _reference_flatline_evidence(
        self,
        report: pd.DataFrame,
        reference_paths: Iterable[Path],
    ) -> pd.DataFrame:
        plants = set(report["plant"].astype(str))
        start = pd.to_datetime(report["start"], errors="coerce").min()
        end = pd.to_datetime(report["end"], errors="coerce").max()
        parts: list[pd.DataFrame] = []
        for path in reference_paths:
            source = pd.read_csv(
                Path(path),
                usecols=["timestamp", "company", "plant", "generation_mwh"],
            )
            source["timestamp"] = pd.to_datetime(source["timestamp"], errors="coerce")
            source = source.loc[
                source["plant"].astype(str).isin(plants)
                & source["timestamp"].between(start, end)
            ]
            if not source.empty:
                parts.append(source)
        if not parts:
            return pd.DataFrame(
                columns=["company", "plant", "reference_rows", "reference_positive_flatline_rate"]
            )
        reference = pd.concat(parts, ignore_index=True)
        reference = reference.groupby(
            ["company", "plant", "timestamp"], as_index=False, sort=False
        )["generation_mwh"].sum()
        rows: list[dict[str, object]] = []
        for (company, plant), group in reference.groupby(["company", "plant"], sort=True):
            rows.append(
                {
                    "company": str(company),
                    "plant": str(plant),
                    "reference_rows": len(group),
                    "reference_positive_flatline_rate": self.profiler._positive_flatline_rate(group),
                }
            )
        return pd.DataFrame(rows)
