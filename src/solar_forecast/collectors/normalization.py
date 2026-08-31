from __future__ import annotations

from pathlib import Path
import re
import warnings
from dataclasses import dataclass

import pandas as pd

from .csv_artifacts import write_standardized_csv


GENERATION_COLUMNS = [
    "timestamp",
    "company",
    "plant_id",
    "plant",
    "unit",
    "energy_source",
    "generation_mwh",
    "capacity_mw",
    "tilt_deg",
    "latitude",
    "longitude",
    "address",
    "source_file",
]

GENERATION_CONTRACT_VERSION = "generation.hourly.v1"


def classify_energy_source(value: object) -> str:
    """Map public Korean generator labels to a stable technology contract."""

    name = str(value).strip().lower()
    if any(token in name for token in ("태양광", "solar", "photovoltaic", "pv")):
        return "solar"
    if any(token in name for token in ("풍력", "wind")):
        return "wind"
    if any(token in name for token in ("수력", "hydro")):
        return "hydro"
    if any(token in name for token in ("ess", "에너지저장", "energy storage")):
        return "storage"
    return "unknown"


def read_csv_with_fallback(path: Path, *, index_col: bool | None = None) -> pd.DataFrame:
    """Read public CSV exports without coupling callers to one Korean encoding."""

    last_error: UnicodeDecodeError | None = None
    for encoding in ("utf-8-sig", "utf-8", "cp949"):
        try:
            return pd.read_csv(path, encoding=encoding, index_col=index_col)
        except UnicodeDecodeError as exc:
            last_error = exc
    raise ValueError(f"Unable to decode CSV: {path}") from last_error


def _empty_static_columns(frame: pd.DataFrame, source_file: str | None) -> pd.DataFrame:
    frame["capacity_mw"] = pd.NA
    frame["tilt_deg"] = pd.NA
    frame["latitude"] = pd.NA
    frame["longitude"] = pd.NA
    frame["address"] = pd.NA
    frame["source_file"] = source_file if source_file else pd.NA
    return frame


class KoenGenerationNormalizer:
    """Convert KOEN's daily wide CSV into one generation observation per hour."""

    required_columns = {"발전구분", "호기", "일자"}
    hour_pattern = re.compile(r"^(\d{1,2})시\s*발전량\((KWh|MWh)\)$", re.IGNORECASE)

    def read(self, path: Path) -> pd.DataFrame:
        # KOEN rows currently include one empty trailing field not declared in
        # the header. index_col=False keeps the leading plant column aligned.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", pd.errors.ParserWarning)
            frame = read_csv_with_fallback(path, index_col=False)
        return self.transform(frame, source_file=path.name)

    def transform(self, frame: pd.DataFrame, *, source_file: str | None = None) -> pd.DataFrame:
        source = frame.copy()
        source.columns = [str(column).strip() for column in source.columns]
        missing = self.required_columns - set(source.columns)
        if missing:
            raise ValueError(f"KOEN columns are missing: {sorted(missing)}")
        hour_matches = {
            column: self.hour_pattern.match(column)
            for column in source.columns
            if self.hour_pattern.match(column)
        }
        hour_columns = list(hour_matches)
        if len(hour_columns) != 24:
            raise ValueError(f"Expected 24 KOEN hourly columns, found {len(hour_columns)}")
        units = {match.group(2).lower() for match in hour_matches.values() if match is not None}
        if len(units) != 1:
            raise ValueError(f"KOEN hourly columns use inconsistent units: {sorted(units)}")
        if source.duplicated(["발전구분", "호기", "일자"]).any():
            raise ValueError("KOEN source contains duplicate plant/unit/date keys")

        long = source.melt(
            id_vars=["발전구분", "호기", "일자"],
            value_vars=hour_columns,
            var_name="source_hour",
            value_name="generation_mwh",
        )
        base_date = pd.to_datetime(long["일자"].astype(str).str.strip(), errors="coerce").dt.normalize()
        hour = long["source_hour"].str.extract(r"^(\d{1,2})", expand=False).astype(int) - 1
        long["timestamp"] = base_date + pd.to_timedelta(hour, unit="h")
        long["plant"] = long["발전구분"].astype(str).str.strip()
        long["unit"] = long["호기"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        long["plant_id"] = "koen:" + long["plant"] + "#" + long["unit"]
        long["generation_mwh"] = pd.to_numeric(long["generation_mwh"], errors="coerce")
        # Older KOEN exports labelled the same kWh-scale values as MWh. The
        # official export changed the header to KWh without changing the value
        # scale. Values above 1,000 MWh in one hour are outside this portfolio's
        # physical scale, so treat that legacy header as a unit-label defect.
        legacy_mislabeled_kwh = units == {"mwh"} and long["generation_mwh"].max(skipna=True) > 1_000
        if units == {"kwh"} or legacy_mislabeled_kwh:
            long["generation_mwh"] = long["generation_mwh"] / 1_000
        long["company"] = "koen"
        long["energy_source"] = "solar"
        long = _empty_static_columns(long, source_file)
        result = long[GENERATION_COLUMNS].dropna(subset=["timestamp", "generation_mwh"])
        result = result.sort_values(["timestamp", "plant_id"], kind="stable").reset_index(drop=True)
        if result.duplicated(["timestamp", "plant_id"]).any():
            raise ValueError("Normalized KOEN data contains duplicate timestamp/plant_id keys")
        declared_unit = next(iter(units))
        result.attrs["generation_unit_resolution"] = {
            "declared_source_unit": declared_unit,
            "resolved_source_unit": (
                "kwh" if declared_unit == "kwh" or legacy_mislabeled_kwh else declared_unit
            ),
            "method": (
                "legacy_header_scale_correction"
                if legacy_mislabeled_kwh
                else "declared_header"
            ),
            "capacity_factor_p99": None,
            "capacity_factor_max": None,
        }
        return result

    def write(self, source: Path, destination: Path) -> Path:
        normalized = self.read(source)
        return write_standardized_csv(normalized, destination)


class EwpTrainingNormalizer:
    """Validate and standardize EWP's nationwide hourly solar training file."""

    column_map = {
        "시도명": "region",
        "설비용량(MW)": "capacity_mw",
        "발전일자": "timestamp",
        "기온": "temperature_c",
        "강우량(mm)": "rainfall_mm",
        "습도": "humidity_pct",
        "적설량(mm)": "snowfall_mm",
        "풍속": "wind_speed_mps",
        "적운량(10분위)": "cloud_amount_tenths",
        "적운량(3분위)": "cloud_amount_thirds",
        "일조(hr)": "sunshine_hours",
        "대기권밖일사량계산값": "extraterrestrial_irradiance",
        "일사량": "solar_irradiance",
        "발전량(MWh)": "generation_mwh",
        "윤년여부": "is_leap_year",
    }
    numeric_columns = [
        "capacity_mw",
        "temperature_c",
        "rainfall_mm",
        "humidity_pct",
        "snowfall_mm",
        "wind_speed_mps",
        "cloud_amount_tenths",
        "cloud_amount_thirds",
        "sunshine_hours",
        "extraterrestrial_irradiance",
        "solar_irradiance",
        "generation_mwh",
    ]

    def read(self, path: Path) -> pd.DataFrame:
        return self.transform(read_csv_with_fallback(path))

    def transform(self, frame: pd.DataFrame) -> pd.DataFrame:
        source = frame.copy()
        source.columns = [str(column).strip() for column in source.columns]
        missing = set(self.column_map) - set(source.columns)
        if missing:
            raise ValueError(f"EWP columns are missing: {sorted(missing)}")
        result = source[list(self.column_map)].rename(columns=self.column_map)
        result["timestamp"] = pd.to_datetime(result["timestamp"], errors="coerce")
        for column in self.numeric_columns:
            result[column] = pd.to_numeric(result[column], errors="coerce")
        result["is_leap_year"] = (
            result["is_leap_year"].astype(str).str.strip().str.upper().map({"Y": 1, "N": 0})
        )
        result["region"] = result["region"].astype(str).str.strip()
        result = result.dropna(subset=["timestamp", "generation_mwh"])
        # The official file overlaps 2024 revisions. The later occurrence is
        # the corrected, higher-precision record for the same region/hour.
        result = result.drop_duplicates(["timestamp", "region"], keep="last")
        result["company"] = "ewp"
        result["hour"] = result["timestamp"].dt.hour
        result["dayofweek"] = result["timestamp"].dt.dayofweek
        result["month"] = result["timestamp"].dt.month
        columns = [
            "timestamp",
            "company",
            "region",
            "capacity_mw",
            "temperature_c",
            "rainfall_mm",
            "humidity_pct",
            "snowfall_mm",
            "wind_speed_mps",
            "cloud_amount_tenths",
            "cloud_amount_thirds",
            "sunshine_hours",
            "extraterrestrial_irradiance",
            "solar_irradiance",
            "is_leap_year",
            "hour",
            "dayofweek",
            "month",
            "generation_mwh",
        ]
        result = result[columns].sort_values(["timestamp", "region"], kind="stable").reset_index(drop=True)
        if result.duplicated(["timestamp", "region"]).any():
            raise ValueError("Normalized EWP data contains duplicate timestamp/region keys")
        return result

    def write(self, source: Path, destination: Path) -> Path:
        normalized = self.read(source)
        return write_standardized_csv(normalized, destination)


@dataclass(frozen=True)
class DailyWideSchema:
    company: str
    date_column: str
    plant_column: str
    hour_columns: tuple[str, ...]
    source_unit: str
    source_unit_alternatives: tuple[str, ...] = ()
    daily_total_column: str | None = None
    daily_total_unit: str | None = None
    daily_total_unit_alternatives: tuple[str, ...] = ()
    unit_column: str | None = None
    id_columns: tuple[str, ...] = ()
    capacity_column: str | None = None
    capacity_unit: str = "mw"
    tilt_column: str | None = None
    latitude_column: str | None = None
    longitude_column: str | None = None
    address_column: str | None = None
    energy_source: str = "solar"
    energy_source_column: str | None = None
    plant_from_filename: bool = False
    include_column: str | None = None
    include_pattern: str | None = None
    duplicate_sequence: bool = False


class DailyWideGenerationNormalizer:
    """Normalize one-row-per-day public generation files to hourly MWh."""

    unit_divisors = {"wh": 1_000_000, "kwh": 1_000, "mwh": 1}
    capacity_divisors = {"w": 1_000_000, "kw": 1_000, "mw": 1}
    extreme_capacity_factor_threshold = 100.0
    plausible_capacity_factor_max = 1.2
    daily_total_tolerance_mwh = 0.001

    def __init__(self, schema: DailyWideSchema):
        if len(schema.hour_columns) != 24:
            raise ValueError("A daily-wide schema must declare exactly 24 hourly columns")
        if schema.source_unit.lower() not in self.unit_divisors:
            raise ValueError(f"Unsupported source energy unit: {schema.source_unit}")
        alternatives = tuple(unit.lower() for unit in schema.source_unit_alternatives)
        unsupported = sorted(set(alternatives) - set(self.unit_divisors))
        if unsupported:
            raise ValueError(f"Unsupported alternative source energy units: {unsupported}")
        if schema.source_unit.lower() in alternatives or len(set(alternatives)) != len(alternatives):
            raise ValueError("Alternative source energy units must be unique and exclude source_unit")
        if alternatives and not schema.capacity_column:
            raise ValueError("Source-unit alternatives require a capacity column for physical validation")
        if bool(schema.daily_total_column) != bool(schema.daily_total_unit):
            raise ValueError("Daily total column and unit must be configured together")
        daily_total_units = tuple(
            unit.lower() for unit in schema.daily_total_unit_alternatives
        )
        if schema.daily_total_unit:
            declared_total_unit = schema.daily_total_unit.lower()
            unsupported_totals = sorted(
                {declared_total_unit, *daily_total_units} - set(self.unit_divisors)
            )
            if unsupported_totals:
                raise ValueError(f"Unsupported daily total energy units: {unsupported_totals}")
            if (
                declared_total_unit in daily_total_units
                or len(set(daily_total_units)) != len(daily_total_units)
            ):
                raise ValueError(
                    "Alternative daily total units must be unique and exclude daily_total_unit"
                )
        elif daily_total_units:
            raise ValueError("Daily total alternatives require a daily total column and unit")
        if schema.capacity_column and schema.capacity_unit.lower() not in self.capacity_divisors:
            raise ValueError(f"Unsupported capacity unit: {schema.capacity_unit}")
        self.schema = schema

    @classmethod
    def _capacity_factor_metrics(
        cls,
        source_values: pd.Series,
        capacity_mw: pd.Series,
        source_unit: str,
    ) -> tuple[float | None, float | None]:
        valid = source_values.notna() & source_values.ge(0) & capacity_mw.notna() & capacity_mw.gt(0)
        if not valid.any():
            return None, None
        ratios = (
            source_values.loc[valid]
            / cls.unit_divisors[source_unit]
            / capacity_mw.loc[valid]
        )
        return float(ratios.quantile(0.99)), float(ratios.max())

    def _resolve_source_unit(
        self,
        source_values: pd.Series,
        capacity_mw: pd.Series,
    ) -> dict[str, object]:
        declared = self.schema.source_unit.lower()
        declared_p99, declared_max = self._capacity_factor_metrics(
            source_values,
            capacity_mw,
            declared,
        )
        resolution: dict[str, object] = {
            "declared_source_unit": declared,
            "resolved_source_unit": declared,
            "method": "declared_header",
            "capacity_factor_p99": declared_p99,
            "capacity_factor_max": declared_max,
        }
        if (
            not self.schema.source_unit_alternatives
            or declared_p99 is None
            or declared_p99 < self.extreme_capacity_factor_threshold
        ):
            return resolution

        plausible: list[tuple[str, float | None, float | None]] = []
        for candidate in self.schema.source_unit_alternatives:
            candidate = candidate.lower()
            candidate_p99, candidate_max = self._capacity_factor_metrics(
                source_values,
                capacity_mw,
                candidate,
            )
            if candidate_max is not None and candidate_max <= self.plausible_capacity_factor_max:
                plausible.append((candidate, candidate_p99, candidate_max))
        if len(plausible) != 1:
            raise ValueError(
                "Declared hourly energy unit is physically impossible, but no unique safe "
                f"alternative was found: declared={declared}, p99_capacity_factor={declared_p99:.3f}"
            )
        resolved, resolved_p99, resolved_max = plausible[0]
        return {
            "declared_source_unit": declared,
            "resolved_source_unit": resolved,
            "method": "capacity_factor_1000x_correction",
            "capacity_factor_p99": resolved_p99,
            "capacity_factor_max": resolved_max,
            "declared_capacity_factor_p99": declared_p99,
            "declared_capacity_factor_max": declared_max,
        }

    def _reconcile_daily_totals(
        self,
        source: pd.DataFrame,
        resolved_hourly_unit: str,
    ) -> dict[str, object] | None:
        total_column = self.schema.daily_total_column
        declared_total_unit = self.schema.daily_total_unit
        if not total_column or not declared_total_unit:
            return None

        hourly = source[list(self.schema.hour_columns)].apply(pd.to_numeric, errors="coerce")
        hourly_source_sum = hourly.sum(axis=1, min_count=len(self.schema.hour_columns))
        hourly_mwh = hourly_source_sum / self.unit_divisors[resolved_hourly_unit]
        daily_total = pd.to_numeric(source[total_column], errors="coerce")
        complete = hourly_mwh.notna() & daily_total.notna()
        incomplete = ~complete
        if incomplete.any():
            sample_indexes = list(source.index[incomplete][:3])
            raise ValueError(
                "Daily-total unit reconciliation requires all 24 hourly buckets and the "
                f"declared total; rows={int(incomplete.sum())}, sample_indexes={sample_indexes}"
            )
        zero = complete & hourly_mwh.eq(0) & daily_total.eq(0)
        evaluable = complete & ~zero
        candidate_units = (
            declared_total_unit.lower(),
            *(unit.lower() for unit in self.schema.daily_total_unit_alternatives),
        )
        errors = pd.DataFrame(
            {
                unit: (hourly_mwh - daily_total / self.unit_divisors[unit]).abs()
                for unit in candidate_units
            },
            index=source.index,
        )
        within_tolerance = errors.le(self.daily_total_tolerance_mwh)
        match_count = within_tolerance.sum(axis=1)
        best_unit = errors.idxmin(axis=1)
        unique_match = evaluable & match_count.eq(1)
        ambiguous = evaluable & match_count.gt(1)
        unreconciled = evaluable & match_count.eq(0)
        if unreconciled.any():
            sample_indexes = list(source.index[unreconciled][:3])
            raise ValueError(
                "Hourly generation does not reconcile to the declared daily total under any "
                f"reviewed unit interpretation; rows={int(unreconciled.sum())}, "
                f"sample_indexes={sample_indexes}"
            )

        date_values = pd.to_datetime(source[self.schema.date_column], errors="coerce")
        resolved_by_row = pd.Series(pd.NA, index=source.index, dtype="string")
        resolved_by_row.loc[evaluable] = best_unit.loc[evaluable]
        if zero.any():
            chronological = pd.DataFrame(
                {"date": date_values, "unit": resolved_by_row},
                index=source.index,
            ).sort_values("date", kind="stable")
            chronological["unit"] = (
                chronological["unit"]
                .ffill()
                .bfill()
                .fillna(declared_total_unit.lower())
            )
            resolved_by_row.loc[zero] = chronological.loc[zero, "unit"]

        evidence_counts: dict[str, int] = {}
        unit_counts: dict[str, int] = {}
        unit_ranges: dict[str, dict[str, str | None]] = {}
        for unit in candidate_units:
            evidence = unique_match & best_unit.eq(unit)
            evidence_count = int(evidence.sum())
            if evidence_count:
                evidence_counts[unit] = evidence_count
            selected = complete & resolved_by_row.eq(unit)
            count = int(selected.sum())
            if not count:
                continue
            unit_counts[unit] = count
            dates = date_values.loc[selected].dropna()
            unit_ranges[unit] = {
                "start": dates.min().date().isoformat() if not dates.empty else None,
                "end": dates.max().date().isoformat() if not dates.empty else None,
            }
        status = "not_evaluable"
        if len(unit_counts) == 1:
            status = "consistent"
        elif len(unit_counts) > 1:
            status = "mixed"
        return {
            "column": total_column,
            "declared_unit": declared_total_unit.lower(),
            "status": status,
            "source_rows": int(len(source)),
            "complete_rows": int(complete.sum()),
            "incomplete_rows": int(incomplete.sum()),
            "zero_rows": int(zero.sum()),
            "ambiguous_rows": int(ambiguous.sum()),
            "unreconciled_rows": int(unreconciled.sum()),
            "evidence_unit_counts": evidence_counts,
            "resolved_unit_counts": unit_counts,
            "resolved_unit_date_ranges": unit_ranges,
            "absolute_tolerance_mwh": self.daily_total_tolerance_mwh,
        }

    def read(self, path: Path) -> pd.DataFrame:
        source = read_csv_with_fallback(path)
        if self.schema.plant_from_filename:
            match = re.search(r"\[([^]]+)]", path.stem)
            if not match:
                raise ValueError(f"Unable to extract plant name from filename: {path.name}")
            source[self.schema.plant_column] = match.group(1).strip()
        return self.transform(source, source_file=path.name)

    def transform(self, frame: pd.DataFrame, *, source_file: str | None = None) -> pd.DataFrame:
        source = frame.copy()
        source.columns = [str(column).strip() for column in source.columns]
        if self.schema.include_column and self.schema.include_pattern:
            if self.schema.include_column not in source.columns:
                raise ValueError(f"{self.schema.company} filter column is missing: {self.schema.include_column}")
            source = source[
                source[self.schema.include_column]
                .astype(str)
                .str.contains(self.schema.include_pattern, regex=True, na=False)
            ].copy()
        # Some official attachments repeat a complete line verbatim. Removing
        # exact copies is deterministic; conflicting rows for the same key are
        # still rejected by the identity check below.
        source = source.drop_duplicates().copy()
        if self.schema.duplicate_sequence:
            source["_source_instance"] = (
                source.groupby([self.schema.date_column, self.schema.plant_column], sort=False)
                .cumcount()
                .add(1)
            )
        identity_columns = [self.schema.date_column, self.schema.plant_column]
        if self.schema.unit_column:
            identity_columns.append(self.schema.unit_column)
        identity_columns.extend(self.schema.id_columns)
        static_columns = [
            column
            for column in (
                self.schema.capacity_column,
                self.schema.tilt_column,
                self.schema.latitude_column,
                self.schema.longitude_column,
                self.schema.address_column,
                self.schema.energy_source_column,
            )
            if column
        ]
        identity_columns = list(dict.fromkeys(identity_columns))
        static_columns = list(dict.fromkeys(static_columns))
        required = set(identity_columns) | set(self.schema.hour_columns) | set(static_columns)
        if self.schema.daily_total_column:
            required.add(self.schema.daily_total_column)
        missing = required - set(source.columns)
        if missing:
            raise ValueError(f"{self.schema.company} columns are missing: {sorted(missing)}")
        if source.duplicated(identity_columns).any():
            raise ValueError(f"{self.schema.company} source contains duplicate daily identity keys")
        long = source.melt(
            id_vars=list(dict.fromkeys([*identity_columns, *static_columns])),
            value_vars=list(self.schema.hour_columns),
            var_name="source_hour",
            value_name="generation_source",
        )
        hour_map = {column: hour - 1 for hour, column in enumerate(self.schema.hour_columns, start=1)}
        base_date = pd.to_datetime(long[self.schema.date_column], errors="coerce").dt.normalize()
        long["timestamp"] = base_date + pd.to_timedelta(long["source_hour"].map(hour_map), unit="h")
        long["plant"] = long[self.schema.plant_column].astype(str).str.strip()
        if self.schema.unit_column:
            long["unit"] = long[self.schema.unit_column].astype(str).str.strip()
        else:
            long["unit"] = ""
        id_parts = [long["plant"]]
        if self.schema.unit_column:
            id_parts.append(long["unit"])
        id_parts.extend(long[column].astype(str).str.strip() for column in self.schema.id_columns)
        long["plant_id"] = self.schema.company + ":" + id_parts[0]
        for part in id_parts[1:]:
            long["plant_id"] = long["plant_id"] + "#" + part
        long["capacity_mw"] = pd.NA
        if self.schema.capacity_column:
            capacity_divisor = self.capacity_divisors[self.schema.capacity_unit.lower()]
            long["capacity_mw"] = (
                pd.to_numeric(long[self.schema.capacity_column], errors="coerce") / capacity_divisor
            )
        source_values = pd.to_numeric(long["generation_source"], errors="coerce")
        unit_resolution = self._resolve_source_unit(source_values, long["capacity_mw"])
        divisor = self.unit_divisors[str(unit_resolution["resolved_source_unit"])]
        daily_total_resolution = self._reconcile_daily_totals(
            source,
            str(unit_resolution["resolved_source_unit"]),
        )
        if daily_total_resolution is not None:
            unit_resolution["daily_total"] = daily_total_resolution
        long["generation_mwh"] = source_values / divisor
        long["company"] = self.schema.company
        source_labels = (
            long[self.schema.energy_source_column]
            if self.schema.energy_source_column
            else long["plant"]
        )
        classified = source_labels.map(classify_energy_source)
        long["energy_source"] = classified.where(
            classified.ne("unknown"), self.schema.energy_source
        )
        for output, source_column in (
            ("tilt_deg", self.schema.tilt_column),
            ("latitude", self.schema.latitude_column),
            ("longitude", self.schema.longitude_column),
        ):
            long[output] = (
                pd.to_numeric(long[source_column], errors="coerce") if source_column else pd.NA
            )
        if self.schema.latitude_column and self.schema.longitude_column:
            # Some Korean public exports label longitude (126~129) as latitude
            # and latitude (33~38) as longitude. Correct only the unambiguous
            # Korea-range inversion and quarantine any remaining invalid pair.
            swapped = long["latitude"].between(124, 132) & long["longitude"].between(32, 39)
            original_latitude = long.loc[swapped, "latitude"].copy()
            long.loc[swapped, "latitude"] = long.loc[swapped, "longitude"].to_numpy()
            long.loc[swapped, "longitude"] = original_latitude.to_numpy()
            valid = long["latitude"].between(32, 39) & long["longitude"].between(124, 132)
            long.loc[~valid, ["latitude", "longitude"]] = pd.NA
        long["address"] = (
            long[self.schema.address_column].astype("string").str.strip()
            if self.schema.address_column
            else pd.NA
        )
        long["source_file"] = source_file if source_file else pd.NA
        result = long[GENERATION_COLUMNS].dropna(subset=["timestamp", "generation_mwh"])
        result = result.sort_values(["timestamp", "plant_id"], kind="stable").reset_index(drop=True)
        if result.duplicated(["timestamp", "plant_id"]).any():
            raise ValueError(f"Normalized {self.schema.company} data contains duplicate hourly keys")
        result.attrs["generation_unit_resolution"] = unit_resolution
        return result

    def write(self, source: Path, destination: Path) -> Path:
        normalized = self.read(source)
        return write_standardized_csv(normalized, destination)


KRC_YEONGAM_SCHEMA = DailyWideSchema(
    company="krc",
    date_column="_date",
    plant_column="시설명",
    hour_columns=tuple(f"{hour}시" for hour in range(1, 25)),
    # The portal labels the buckets as kW, while the file sums the 24 hourly
    # buckets into a daily total. For a one-hour bucket the numeric value is
    # kWh-equivalent, so the explicit conversion to hourly MWh is / 1,000.
    source_unit="kwh",
    capacity_column="_capacity_mw",
    capacity_unit="mw",
    address_column="_address",
)


class KrcYeongamGenerationNormalizer:
    """Normalize the Korean Rural Community Corporation Yeongam exports.

    The public series changed its hour headings over time and did not include
    a plant identifier before 2022. Files without ``시설명`` are rejected
    instead of guessing which physical asset generated the observations.
    """

    capacity_mw = {
        "영암1차": 1.4916,
        "영암2차": 1.4916,
        "율치": 0.9198,
    }
    hour_pattern = re.compile(r"^(\d{1,2})(?:시(?:\(h\))?|h)$", re.IGNORECASE)
    year_pattern = re.compile(r"(20\d{2})(?:1231)?")

    def read(self, path: Path) -> pd.DataFrame:
        match = self.year_pattern.search(path.stem)
        if not match:
            raise ValueError(f"Unable to infer KRC observation year from filename: {path.name}")
        return self.transform(
            read_csv_with_fallback(path),
            year=int(match.group(1)),
            source_file=path.name,
        )

    def transform(
        self,
        frame: pd.DataFrame,
        *,
        year: int,
        source_file: str | None = None,
    ) -> pd.DataFrame:
        source = frame.copy()
        source.columns = [str(column).strip() for column in source.columns]
        renamed: dict[str, str] = {}
        for column in source.columns:
            compact = re.sub(r"\s+", "", column)
            match = self.hour_pattern.match(compact)
            if match:
                renamed[column] = f"{int(match.group(1))}시"
        source = source.rename(columns=renamed)

        required = {"월", "일", "시설명", *KRC_YEONGAM_SCHEMA.hour_columns}
        missing = required - set(source.columns)
        if missing:
            raise ValueError(
                "KRC Yeongam file is not entity-safe; required columns are missing: "
                f"{sorted(missing)}"
            )
        source["시설명"] = source["시설명"].astype("string").str.strip()
        unknown = sorted(set(source["시설명"].dropna()) - set(self.capacity_mw))
        if unknown:
            raise ValueError(f"Unknown KRC Yeongam facilities require metadata review: {unknown}")

        source["_date"] = pd.to_datetime(
            {
                "year": year,
                "month": pd.to_numeric(source["월"], errors="coerce"),
                "day": pd.to_numeric(source["일"], errors="coerce"),
            },
            errors="coerce",
        )
        if source["_date"].isna().any():
            raise ValueError("KRC Yeongam file contains invalid month/day values")
        if source.duplicated(["_date", "시설명"]).any():
            raise ValueError("KRC Yeongam file contains duplicate facility/date keys")

        hourly = source[list(KRC_YEONGAM_SCHEMA.hour_columns)].apply(
            pd.to_numeric, errors="coerce"
        )
        if "계" in source:
            declared_total = pd.to_numeric(source["계"], errors="coerce")
            complete = hourly.notna().all(axis=1) & declared_total.notna()
            difference = hourly.loc[complete].sum(axis=1).sub(declared_total.loc[complete]).abs()
            if not difference.empty and difference.max() > 1e-6:
                raise ValueError("KRC Yeongam hourly buckets do not reconcile to the daily total")
        source[list(KRC_YEONGAM_SCHEMA.hour_columns)] = hourly
        source["_capacity_mw"] = source["시설명"].map(self.capacity_mw)
        source["_address"] = "전라남도 영암군"
        return DailyWideGenerationNormalizer(KRC_YEONGAM_SCHEMA).transform(
            source,
            source_file=source_file,
        )


KOSPO_WIDE_SCHEMA = DailyWideSchema(
    company="kospo",
    date_column="거래일자",
    plant_column="발전소명",
    unit_column="발전기명",
    id_columns=("계량구분",),
    hour_columns=tuple(f"{hour}시" for hour in range(1, 25)),
    source_unit="kwh",
)

IWEST_WIDE_SCHEMA = DailyWideSchema(
    company="iwest",
    date_column="년월일",
    plant_column="발전기명",
    hour_columns=tuple(f"{hour:02d}시" for hour in range(1, 25)),
    source_unit="wh",
    capacity_column="설비용량(MW)",
    capacity_unit="mw",
)


KOSPO_ARCHIVE_LEGACY_SCHEMA = DailyWideSchema(
    company="kospo",
    date_column="년월일",
    plant_column="_source_plant",
    unit_column="호기",
    hour_columns=tuple(str(hour) for hour in range(1, 25)),
    source_unit="kwh",
    plant_from_filename=True,
)

KOSPO_ARCHIVE_INTERVAL_SCHEMA = DailyWideSchema(
    company="kospo",
    date_column="년월일",
    plant_column="_source_plant",
    unit_column="호기",
    hour_columns=tuple(f"{hour}-{hour + 1}시 발전량(kWh)" for hour in range(24)),
    source_unit="kwh",
    # One official export variant keeps the kWh header but emits hourly values
    # at Wh scale. Resolve that documented ambiguity only when the declared
    # interpretation is over 100x capacity and the Wh interpretation remains
    # within a conservative one-hour physical bound.
    source_unit_alternatives=("wh",),
    daily_total_column="총량(kWh)",
    daily_total_unit="kwh",
    # The Iksan Dasong export changes this column from kWh-scale totals to
    # Wh-scale totals on 2024-10-01 without changing its header. The hourly
    # buckets remain Wh-scale on both sides of that boundary.
    daily_total_unit_alternatives=("wh",),
    capacity_column="설비용량(kW)",
    capacity_unit="kw",
    plant_from_filename=True,
)

EWP_WH_SCHEMA = DailyWideSchema(
    company="ewp",
    date_column="날짜",
    plant_column="발전기명",
    hour_columns=("01시(Wh)", *(f"{hour:02d}시" for hour in range(2, 25))),
    source_unit="wh",
)

EWP_POINT_SCHEMA = DailyWideSchema(
    company="ewp",
    date_column="날짜",
    plant_column="발전기명",
    hour_columns=("01시(kWh)", *(f"{hour:02d}시" for hour in range(2, 25))),
    # The official heading says kWh, but the values are Wh-scale: for example
    # a 0.7 MW plant reports about 551,904 at peak. Dividing by 1e6 produces
    # the physically valid 0.552 MWh; treating it as kWh would yield 552 MWh.
    source_unit="wh",
    capacity_column="설비용량(메가와트)",
    capacity_unit="mw",
    latitude_column="위도",
    longitude_column="경도",
)

IWEST_RENEWABLE_SCHEMA = DailyWideSchema(
    company="iwest",
    date_column="날짜",
    plant_column="발전기명",
    hour_columns=tuple(f"{hour:02d}시" for hour in range(1, 25)),
    source_unit="kwh",
    capacity_column="용량_메가와트",
    capacity_unit="mw",
    id_columns=("_source_instance",),
    energy_source_column="발전기명",
    duplicate_sequence=True,
)

IWEST_ARCHIVE_STATUS_SCHEMA = DailyWideSchema(
    company="iwest",
    date_column="년월일",
    plant_column="발전기명",
    hour_columns=("01시(Wh)", *(f"{hour:02d}시" for hour in range(2, 25))),
    source_unit="wh",
    capacity_column="설비용량(MW)",
    capacity_unit="mw",
)


ARCHIVE_WIDE_SCHEMAS = (
    KOSPO_ARCHIVE_INTERVAL_SCHEMA,
    KOSPO_ARCHIVE_LEGACY_SCHEMA,
    EWP_POINT_SCHEMA,
    EWP_WH_SCHEMA,
    IWEST_ARCHIVE_STATUS_SCHEMA,
    IWEST_RENEWABLE_SCHEMA,
)
