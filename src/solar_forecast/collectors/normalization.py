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

    def __init__(self, schema: DailyWideSchema):
        if len(schema.hour_columns) != 24:
            raise ValueError("A daily-wide schema must declare exactly 24 hourly columns")
        if schema.source_unit.lower() not in self.unit_divisors:
            raise ValueError(f"Unsupported source energy unit: {schema.source_unit}")
        if schema.capacity_column and schema.capacity_unit.lower() not in self.capacity_divisors:
            raise ValueError(f"Unsupported capacity unit: {schema.capacity_unit}")
        self.schema = schema

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
        divisor = self.unit_divisors[self.schema.source_unit.lower()]
        long["generation_mwh"] = pd.to_numeric(long["generation_source"], errors="coerce") / divisor
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
        long["capacity_mw"] = pd.NA
        if self.schema.capacity_column:
            capacity_divisor = self.capacity_divisors[self.schema.capacity_unit.lower()]
            long["capacity_mw"] = (
                pd.to_numeric(long[self.schema.capacity_column], errors="coerce") / capacity_divisor
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
