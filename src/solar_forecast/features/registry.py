from __future__ import annotations

from dataclasses import dataclass
import json
from math import asin, cos, radians, sin, sqrt
from pathlib import Path
import re
from typing import Iterable

import pandas as pd

from solar_forecast.collectors.metadata import PlantMetadataCatalog
from solar_forecast.collectors.normalization import read_csv_with_fallback


PROVINCE_ALIASES = {
    "서울특별시": ("서울특별시", "서울"),
    "부산광역시": ("부산광역시", "부산"),
    "대구광역시": ("대구광역시", "대구"),
    "인천광역시": ("인천광역시", "인천"),
    "광주광역시": ("광주광역시", "광주"),
    "대전광역시": ("대전광역시", "대전"),
    "울산광역시": ("울산광역시", "울산"),
    "세종특별자치시": ("세종특별자치시", "세종시", "세종"),
    "경기도": ("경기도", "경기"),
    "강원특별자치도": ("강원특별자치도", "강원도", "강원"),
    "충청북도": ("충청북도", "충북"),
    "충청남도": ("충청남도", "충남"),
    "전북특별자치도": ("전북특별자치도", "전라북도", "전북"),
    "전라남도": ("전라남도", "전남"),
    "경상북도": ("경상북도", "경북"),
    "경상남도": ("경상남도", "경남"),
    "제주특별자치도": ("제주특별자치도", "제주도", "제주"),
}

METROPOLITAN_PROVINCES = {
    "서울특별시",
    "부산광역시",
    "대구광역시",
    "인천광역시",
    "광주광역시",
    "대전광역시",
    "울산광역시",
    "세종특별자치시",
}

# Some public plant tables omit the province. Keep only municipalities that
# appear in retained official sources; unknown places remain unresolved.
MUNICIPALITY_PROVINCE_OVERRIDES = {
    "제주시": "제주특별자치도",
    "서귀포시": "제주특별자치도",
    "하동군": "경상남도",
}

ADDRESS_ONLY_REVIEW_MUNICIPALITIES = {"옹진군"}


def _normalized_mapping_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.casefold() in {"", "none", "null", "nan"} else text


@dataclass(frozen=True)
class AdministrativeArea:
    province: str | None
    city: str | None
    locality: str | None


@dataclass(frozen=True)
class StationMatch:
    station_id: int | None
    station_name: str | None
    distance_km: float | None
    method: str
    confidence: str
    review_required: bool
    reason: str | None = None


@dataclass(frozen=True)
class ReviewedStationMapping:
    company: str
    plant: str
    station_id: int
    evidence_url: str
    rationale: str

    def __post_init__(self) -> None:
        company = _normalized_mapping_text(self.company)
        plant = _normalized_mapping_text(self.plant)
        evidence_url = _normalized_mapping_text(self.evidence_url)
        rationale = _normalized_mapping_text(self.rationale)
        if not company or not plant:
            raise ValueError("Reviewed weather mapping requires company and plant")
        if not evidence_url or not rationale:
            raise ValueError(
                "Reviewed weather mapping requires evidence_url and rationale"
            )
        try:
            station_id = int(self.station_id)
        except (TypeError, ValueError) as error:
            raise ValueError(
                "Reviewed weather mapping requires a positive station_id"
            ) from error
        if station_id <= 0:
            raise ValueError("Reviewed weather mapping requires a positive station_id")
        object.__setattr__(self, "company", company)
        object.__setattr__(self, "plant", plant)
        object.__setattr__(self, "station_id", station_id)
        object.__setattr__(self, "evidence_url", evidence_url)
        object.__setattr__(self, "rationale", rationale)


class ReviewedStationMappingCatalog:
    """Version-controlled, auditable exceptions for plants without local ASOS."""

    def __init__(self, mappings: Iterable[ReviewedStationMapping] = ()):
        self._mappings: dict[tuple[str, str], ReviewedStationMapping] = {}
        for mapping in mappings:
            key = (mapping.company, mapping.plant)
            if key in self._mappings:
                raise ValueError(f"Duplicate reviewed weather mapping: {key}")
            self._mappings[key] = mapping

    @classmethod
    def from_json(cls, path: Path) -> "ReviewedStationMappingCatalog":
        source = Path(path)
        if not source.exists():
            return cls()
        payload = json.loads(source.read_text(encoding="utf-8"))
        return cls(
            ReviewedStationMapping(
                company=row.get("company"),
                plant=row.get("plant"),
                station_id=row.get("station_id"),
                evidence_url=row.get("evidence_url"),
                rationale=row.get("rationale"),
            )
            for row in payload.get("mappings", [])
        )

    def get(self, company: str, plant: str) -> ReviewedStationMapping | None:
        return self._mappings.get((str(company), str(plant)))


def parse_administrative_area(address: object) -> AdministrativeArea:
    if address is None or pd.isna(address):
        return AdministrativeArea(None, None, None)
    text = re.sub(r"\s+", " ", str(address).strip())
    province = next(
        (
            canonical
            for canonical, aliases in PROVINCE_ALIASES.items()
            if any(re.search(rf"(^|\s){re.escape(alias)}(?=\s|$)", text) for alias in aliases)
        ),
        None,
    )
    subdivisions = re.findall(r"[가-힣]+(?:시|군|구)", text)
    city = None
    for subdivision in subdivisions:
        if subdivision in {alias for aliases in PROVINCE_ALIASES.values() for alias in aliases}:
            continue
        city = subdivision
        break
    if province is None and city in MUNICIPALITY_PROVINCE_OVERRIDES:
        province = MUNICIPALITY_PROVINCE_OVERRIDES[city]
    if province in METROPOLITAN_PROVINCES and city is None:
        city = province
    localities = re.findall(r"[가-힣0-9]+(?:읍|면|동|리)", text)
    return AdministrativeArea(province, city, localities[0] if localities else None)


def _haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    earth_radius_km = 6371.0088
    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)
    a = sin(dlat / 2) ** 2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon / 2) ** 2
    return 2 * earth_radius_km * asin(sqrt(a))


class KmaStationCatalog:
    """Resolve a plant to an ASOS station without hiding ambiguous guesses."""

    def __init__(self, stations: pd.DataFrame):
        self.stations = stations.reset_index(drop=True)

    @classmethod
    def from_metadata(cls, path: Path) -> "KmaStationCatalog":
        source = read_csv_with_fallback(Path(path))
        source.columns = [str(column).strip() for column in source.columns]
        required = {"지점", "지점명", "지점주소", "위도", "경도"}
        missing = required - set(source.columns)
        if missing:
            raise ValueError(f"KMA station metadata columns are missing: {sorted(missing)}")
        for column in ("시작일", "종료일"):
            if column not in source:
                source[column] = pd.NaT
            source[column] = pd.to_datetime(source[column], errors="coerce")
        source = source.sort_values(["지점", "시작일"], kind="stable").drop_duplicates(
            ["지점", "시작일", "종료일"], keep="last"
        )
        rows: list[dict[str, object]] = []
        for values in source.to_dict("records"):
            area = parse_administrative_area(values["지점주소"])
            rows.append(
                {
                    "station_id": int(values["지점"]),
                    "station_name": str(values["지점명"]).strip(),
                    "station_address": str(values["지점주소"]).strip(),
                    "station_latitude": pd.to_numeric(values["위도"], errors="coerce"),
                    "station_longitude": pd.to_numeric(values["경도"], errors="coerce"),
                    "station_valid_from": values["시작일"],
                    "station_valid_to": values["종료일"],
                    "admin_province": area.province,
                    "admin_city": area.city,
                    "admin_locality": area.locality,
                }
            )
        return cls(pd.DataFrame(rows))

    def by_id(
        self,
        station_id: int,
        *,
        generation_start: object = None,
        generation_end: object = None,
    ) -> pd.Series | None:
        rows = self.stations.loc[self.stations["station_id"].eq(int(station_id))]
        rows = self._covering_stations(
            rows,
            generation_start=generation_start,
            generation_end=generation_end,
        )
        if rows.empty:
            return None
        if "station_valid_from" in rows:
            rows = rows.sort_values("station_valid_from", kind="stable")
        return rows.iloc[-1]

    def resolve(
        self,
        *,
        address: object,
        latitude: object,
        longitude: object,
        reviewed_station_id: int | None = None,
        generation_start: object = None,
        generation_end: object = None,
    ) -> StationMatch:
        if reviewed_station_id is not None:
            station = self.by_id(
                reviewed_station_id,
                generation_start=generation_start,
                generation_end=generation_end,
            )
            if station is not None:
                return StationMatch(
                    int(station["station_id"]),
                    str(station["station_name"]),
                    self._distance(latitude, longitude, station),
                    "reviewed_config_mapping",
                    "high",
                    False,
                )
            return StationMatch(
                None,
                None,
                None,
                "reviewed_mapping_invalid",
                "none",
                True,
                "reviewed ASOS station does not exist or cover the complete generation date range",
            )

        lat = pd.to_numeric(latitude, errors="coerce")
        lon = pd.to_numeric(longitude, errors="coerce")
        if pd.notna(lat) and pd.notna(lon) and 124 <= lat <= 132 and 32 <= lon <= 39:
            lat, lon = lon, lat
        if pd.notna(lat) and pd.notna(lon):
            histories = self._covering_stations(
                self.stations,
                generation_start=generation_start,
                generation_end=generation_end,
            )
            coordinate_rows = histories[
                ["station_latitude", "station_longitude"]
            ].notna().all(axis=1)
            coordinate_complete = coordinate_rows.groupby(
                histories["station_id"], sort=False
            ).all()
            histories = histories.loc[
                histories["station_id"].isin(
                    coordinate_complete.index[coordinate_complete]
                )
            ].copy()
            if histories.empty:
                return StationMatch(
                    None,
                    None,
                    None,
                    "unresolved",
                    "none",
                    True,
                    "no ASOS station covers the complete generation date range",
                )
            histories["distance_km"] = histories.apply(
                lambda row: _haversine_km(
                    float(lat),
                    float(lon),
                    float(row["station_latitude"]),
                    float(row["station_longitude"]),
                ),
                axis=1,
            )
            worst_distance = histories.groupby("station_id", sort=False)[
                "distance_km"
            ].max()
            candidates = self._latest_station_rows(histories)
            candidates["distance_km"] = candidates["station_id"].map(
                worst_distance
            )
            station = candidates.sort_values("distance_km", kind="stable").iloc[0]
            distance = float(station["distance_km"])
            if distance <= 50:
                return StationMatch(
                    int(station["station_id"]),
                    str(station["station_name"]),
                    distance,
                    "coordinate_nearest",
                    "high",
                    False,
                )
            return StationMatch(
                int(station["station_id"]) if distance <= 100 else None,
                str(station["station_name"]) if distance <= 100 else None,
                distance,
                "coordinate_nearest_review" if distance <= 100 else "unresolved",
                "review" if distance <= 100 else "none",
                True,
                "nearest ASOS station is farther than 50 km",
            )

        area = parse_administrative_area(address)
        if area.province and area.city:
            active_stations = self._covering_stations(
                self.stations,
                generation_start=generation_start,
                generation_end=generation_end,
            )
            exact_rows = (
                active_stations["admin_province"].eq(area.province)
                & active_stations["admin_city"].eq(area.city)
            )
            exact_ids = exact_rows.groupby(
                active_stations["station_id"], sort=False
            ).all()
            candidates = active_stations.loc[
                active_stations["station_id"].isin(exact_ids.index[exact_ids])
            ]
            if len(candidates) > 1 and area.locality:
                locality_rows = candidates["admin_locality"].eq(area.locality)
                locality_ids = locality_rows.groupby(
                    candidates["station_id"], sort=False
                ).all()
                locality = candidates.loc[
                    candidates["station_id"].isin(
                        locality_ids.index[locality_ids]
                    )
                ]
                if locality["station_id"].nunique() == 1:
                    candidates = locality
            candidates = self._latest_station_rows(candidates)
            if len(candidates) == 1 and area.city not in ADDRESS_ONLY_REVIEW_MUNICIPALITIES:
                station = candidates.iloc[0]
                return StationMatch(
                    int(station["station_id"]),
                    str(station["station_name"]),
                    None,
                    "administrative_area_exact",
                    "high",
                    False,
                )
            if len(candidates) > 1:
                return StationMatch(
                    None,
                    None,
                    None,
                    "unresolved",
                    "none",
                    True,
                    "multiple ASOS stations match the public address",
                )
        return StationMatch(
            None,
            None,
            None,
            "unresolved",
            "none",
            True,
            "public metadata is insufficient for a defensible ASOS mapping",
        )

    @classmethod
    def _covering_stations(
        cls,
        stations: pd.DataFrame,
        *,
        generation_start: object,
        generation_end: object,
    ) -> pd.DataFrame:
        if stations.empty or "station_valid_from" not in stations:
            return stations
        start = pd.to_datetime(generation_start, errors="coerce")
        end = pd.to_datetime(generation_end, errors="coerce")
        if pd.isna(start) and pd.isna(end):
            return stations
        valid_from = pd.to_datetime(stations["station_valid_from"], errors="coerce")
        valid_to = (
            pd.to_datetime(stations["station_valid_to"], errors="coerce")
            if "station_valid_to" in stations
            else pd.Series(pd.NaT, index=stations.index, dtype="datetime64[ns]")
        )
        mask = pd.Series(True, index=stations.index)
        earliest = start if pd.notna(start) else end
        latest = end if pd.notna(end) else start
        coverage_ids = [
            station_id
            for station_id, indexes in stations.groupby(
                "station_id", sort=False
            ).groups.items()
            if cls._intervals_cover_range(
                valid_from.loc[indexes],
                valid_to.loc[indexes],
                earliest,
                latest,
            )
        ]
        mask &= stations["station_id"].isin(coverage_ids)
        if pd.notna(latest):
            mask &= valid_from.isna() | valid_from.le(latest)
        if pd.notna(earliest):
            mask &= valid_to.isna() | valid_to.ge(earliest)
        return stations.loc[mask]

    @staticmethod
    def _intervals_cover_range(
        valid_from: pd.Series,
        valid_to: pd.Series,
        earliest: pd.Timestamp,
        latest: pd.Timestamp,
    ) -> bool:
        intervals = sorted(
            zip(valid_from.tolist(), valid_to.tolist()),
            key=lambda interval: pd.Timestamp.min
            if pd.isna(interval[0])
            else interval[0],
        )
        cursor = earliest
        for interval_start, interval_end in intervals:
            interval_start = (
                pd.Timestamp.min if pd.isna(interval_start) else interval_start
            )
            interval_end = pd.Timestamp.max if pd.isna(interval_end) else interval_end
            if interval_end < cursor:
                continue
            if interval_start > cursor + pd.Timedelta(days=1):
                return False
            if interval_end >= latest:
                return True
            cursor = max(cursor, interval_end)
        return False

    @staticmethod
    def _latest_station_rows(stations: pd.DataFrame) -> pd.DataFrame:
        if stations.empty:
            return stations
        if "station_valid_from" not in stations:
            return stations.drop_duplicates("station_id", keep="last")
        return stations.sort_values(
            "station_valid_from", kind="stable"
        ).drop_duplicates("station_id", keep="last")

    @staticmethod
    def _distance(latitude: object, longitude: object, station: pd.Series) -> float | None:
        values = [
            pd.to_numeric(latitude, errors="coerce"),
            pd.to_numeric(longitude, errors="coerce"),
            pd.to_numeric(station["station_latitude"], errors="coerce"),
            pd.to_numeric(station["station_longitude"], errors="coerce"),
        ]
        if any(pd.isna(value) for value in values):
            return None
        return _haversine_km(*map(float, values))


class NationwidePlantRegistryBuilder:
    """Create one auditable nationwide identity and station-mapping table."""

    scan_columns = [
        "timestamp",
        "company",
        "plant",
        "energy_source",
        "capacity_mw",
        "tilt_deg",
        "latitude",
        "longitude",
        "address",
    ]

    def __init__(
        self,
        metadata: PlantMetadataCatalog,
        stations: KmaStationCatalog,
        reviewed_mappings: ReviewedStationMappingCatalog | None = None,
    ):
        self.metadata = metadata
        self.stations = stations
        self.reviewed_mappings = reviewed_mappings or ReviewedStationMappingCatalog()

    def build(
        self,
        paths: Iterable[Path],
        destination: Path,
        *,
        legacy_mapping: pd.DataFrame | None = None,
    ) -> pd.DataFrame:
        profiles = self._scan(paths)
        legacy = self._legacy_map(legacy_mapping)
        enriched_profiles: list[dict[str, object]] = []
        for profile in profiles:
            company = str(profile["company"])
            plant = str(profile["plant"])
            energy_source = str(profile["energy_source"])
            record = self.metadata.lookup(
                company,
                plant,
                energy_source=energy_source,
                aggregate=True,
            )
            capacity = record.capacity_mw if record and record.capacity_mw is not None else profile["capacity_mw"]
            tilt = record.tilt_deg if record and record.tilt_deg is not None else profile["tilt_deg"]
            address = record.address if record and record.address else profile["address"]
            latitude = profile["latitude"]
            longitude = profile["longitude"]
            if (
                pd.notna(latitude)
                and pd.notna(longitude)
                and 124 <= float(latitude) <= 132
                and 32 <= float(longitude) <= 39
            ):
                latitude, longitude = longitude, latitude
            area = parse_administrative_area(address)
            enriched_profiles.append(
                {
                    **profile,
                    "capacity_mw": capacity,
                    "tilt_deg": tilt,
                    "address": address,
                    "latitude": latitude,
                    "longitude": longitude,
                    "area": area,
                }
            )

        rows: list[dict[str, object]] = []
        for profile in enriched_profiles:
            company = str(profile["company"])
            plant = str(profile["plant"])
            energy_source = str(profile["energy_source"])
            capacity = profile["capacity_mw"]
            tilt = profile["tilt_deg"]
            address = profile["address"]
            latitude = profile["latitude"]
            longitude = profile["longitude"]
            area = profile["area"]
            reviewed = self.reviewed_mappings.get(company, plant)
            legacy_candidate_id = legacy.get((company, plant))
            legacy_station = (
                self.stations.by_id(legacy_candidate_id)
                if legacy_candidate_id is not None
                else None
            )
            match = self.stations.resolve(
                address=address,
                latitude=latitude,
                longitude=longitude,
                reviewed_station_id=reviewed.station_id if reviewed else None,
                generation_start=profile["generation_start"],
                generation_end=profile["generation_end"],
            )
            eligible = match.station_id is not None and not match.review_required
            model_ready_reason = None if eligible else match.reason
            if not eligible and legacy_candidate_id is not None:
                model_ready_reason = (
                    f"{model_ready_reason}; " if model_ready_reason else ""
                ) + "legacy ASOS station is retained only as an unreviewed audit candidate"
            rows.append(
                {
                    "plant_id": f"{company}:{plant}",
                    "company": company,
                    "plant": plant,
                    "energy_source": energy_source,
                    "capacity_mw": capacity,
                    "tilt_deg": tilt,
                    "address": address,
                    "admin_province": area.province,
                    "admin_city": area.city,
                    "latitude": latitude,
                    "longitude": longitude,
                    "weather_station_id": match.station_id,
                    "weather_station_name": match.station_name,
                    "weather_station_distance_km": match.distance_km,
                    "weather_mapping_method": match.method,
                    "weather_mapping_confidence": match.confidence,
                    "weather_mapping_review_required": match.review_required,
                    "weather_mapping_evidence_url": reviewed.evidence_url if reviewed else None,
                    "weather_mapping_rationale": reviewed.rationale if reviewed else None,
                    "legacy_weather_station_candidate_id": legacy_candidate_id,
                    "legacy_weather_station_candidate_name": (
                        str(legacy_station["station_name"])
                        if legacy_station is not None
                        else None
                    ),
                    "legacy_weather_station_candidate_distance_km": (
                        self.stations._distance(latitude, longitude, legacy_station)
                        if legacy_station is not None
                        else None
                    ),
                    "legacy_weather_station_candidate_status": (
                        "audit_only_unreviewed"
                        if legacy_candidate_id is not None
                        else None
                    ),
                    "generation_start": profile["generation_start"],
                    "generation_end": profile["generation_end"],
                    "source_plant_aliases": profile["source_plant_aliases"],
                    "source_observation_rows": profile["source_observation_rows"],
                    "source_partition_count": profile["source_partition_count"],
                    "model_ready_status": "eligible" if eligible else "quarantined",
                    "model_ready_reason": model_ready_reason,
                }
            )
        result = pd.DataFrame(rows).sort_values(
            ["company", "plant", "energy_source"], kind="stable"
        ).reset_index(drop=True)
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        result.to_csv(destination, index=False, encoding="utf-8-sig")
        return result

    def _scan(self, paths: Iterable[Path]) -> list[dict[str, object]]:
        profiles: dict[tuple[str, str, str], dict[str, object]] = {}
        for path in paths:
            source = pd.read_csv(Path(path), usecols=self.scan_columns)
            source["timestamp"] = pd.to_datetime(source["timestamp"], errors="coerce")
            for key, group in source.groupby(["company", "plant", "energy_source"], sort=False):
                company, source_plant, energy_source = map(str, key)
                plant = self.metadata.canonical_plant(company, source_plant)
                profile = profiles.setdefault(
                    (company, plant, energy_source),
                    {
                        "company": company,
                        "plant": plant,
                        "energy_source": energy_source,
                        "capacity_mw": None,
                        "tilt_deg": None,
                        "latitude": None,
                        "longitude": None,
                        "address": None,
                        "generation_start": None,
                        "generation_end": None,
                        "source_observation_rows": 0,
                        "source_partitions": set(),
                        "source_plant_aliases": set(),
                    },
                )
                for column in ("capacity_mw", "tilt_deg", "latitude", "longitude", "address"):
                    values = group[column].dropna()
                    if profile[column] is None and not values.empty:
                        profile[column] = values.iloc[0]
                start, end = group["timestamp"].min(), group["timestamp"].max()
                if pd.notna(start) and (profile["generation_start"] is None or start < profile["generation_start"]):
                    profile["generation_start"] = start
                if pd.notna(end) and (profile["generation_end"] is None or end > profile["generation_end"]):
                    profile["generation_end"] = end
                profile["source_observation_rows"] = int(profile["source_observation_rows"]) + len(group)
                profile["source_partitions"].add(str(path))
                profile["source_plant_aliases"].add(source_plant)
        result = []
        for profile in profiles.values():
            profile["generation_start"] = profile["generation_start"].isoformat()
            profile["generation_end"] = profile["generation_end"].isoformat()
            profile["source_partition_count"] = len(profile.pop("source_partitions"))
            profile["source_plant_aliases"] = " | ".join(sorted(profile["source_plant_aliases"]))
            result.append(profile)
        return result

    def _legacy_map(self, frame: pd.DataFrame | None) -> dict[tuple[str, str], int]:
        if frame is None or frame.empty:
            return {}
        required = {"company", "plant", "station_id"}
        if not required.issubset(frame.columns):
            return {}
        mapping = frame[list(required)].dropna().drop_duplicates().copy()
        mapping["plant"] = [
            self.metadata.canonical_plant(company, plant)
            for company, plant in mapping[["company", "plant"]].itertuples(index=False, name=None)
        ]
        ambiguous = mapping.groupby(["company", "plant"]).size()
        ambiguous = ambiguous[ambiguous.gt(1)]
        if not ambiguous.empty:
            raise ValueError(f"Legacy source has ambiguous plant/station mappings: {ambiguous.to_dict()}")
        return {
            (str(row.company), str(row.plant)): int(row.station_id)
            for row in mapping.itertuples(index=False)
        }
