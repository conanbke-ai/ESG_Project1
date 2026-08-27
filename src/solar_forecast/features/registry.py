from __future__ import annotations

from dataclasses import dataclass
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
        if "시작일" in source:
            source["시작일"] = pd.to_datetime(source["시작일"], errors="coerce")
            source = source.sort_values("시작일", kind="stable")
        source = source.drop_duplicates("지점", keep="last")
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
                    "admin_province": area.province,
                    "admin_city": area.city,
                    "admin_locality": area.locality,
                }
            )
        return cls(pd.DataFrame(rows))

    def by_id(self, station_id: int) -> pd.Series | None:
        rows = self.stations.loc[self.stations["station_id"].eq(int(station_id))]
        return rows.iloc[0] if len(rows) == 1 else None

    def resolve(
        self,
        *,
        address: object,
        latitude: object,
        longitude: object,
        legacy_station_id: int | None = None,
        reviewed_mapping_method: str = "reviewed_legacy_mapping",
    ) -> StationMatch:
        if legacy_station_id is not None:
            station = self.by_id(legacy_station_id)
            if station is not None:
                return StationMatch(
                    int(station["station_id"]),
                    str(station["station_name"]),
                    self._distance(latitude, longitude, station),
                    reviewed_mapping_method,
                    "high",
                    False,
                )

        lat = pd.to_numeric(latitude, errors="coerce")
        lon = pd.to_numeric(longitude, errors="coerce")
        if pd.notna(lat) and pd.notna(lon) and 124 <= lat <= 132 and 32 <= lon <= 39:
            lat, lon = lon, lat
        if pd.notna(lat) and pd.notna(lon):
            candidates = self.stations.dropna(subset=["station_latitude", "station_longitude"]).copy()
            candidates["distance_km"] = candidates.apply(
                lambda row: _haversine_km(
                    float(lat),
                    float(lon),
                    float(row["station_latitude"]),
                    float(row["station_longitude"]),
                ),
                axis=1,
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
            candidates = self.stations.loc[
                self.stations["admin_province"].eq(area.province)
                & self.stations["admin_city"].eq(area.city)
            ]
            if len(candidates) > 1 and area.locality:
                locality = candidates.loc[candidates["admin_locality"].eq(area.locality)]
                if len(locality) == 1:
                    candidates = locality
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

    def __init__(self, metadata: PlantMetadataCatalog, stations: KmaStationCatalog):
        self.metadata = metadata
        self.stations = stations

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
            record = self.metadata.lookup(company, plant, aggregate=True)
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

        colocated: dict[str, set[int]] = {}
        for profile in enriched_profiles:
            station_id = legacy.get((str(profile["company"]), str(profile["plant"])))
            address_key = self._address_key(profile["address"])
            if station_id is not None and address_key:
                colocated.setdefault(address_key, set()).add(station_id)

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
            legacy_station_id = legacy.get((company, plant))
            reviewed_method = "reviewed_legacy_mapping"
            if legacy_station_id is None:
                colocated_ids = colocated.get(self._address_key(address), set())
                if len(colocated_ids) == 1:
                    legacy_station_id = next(iter(colocated_ids))
                    reviewed_method = "reviewed_colocated_address"
            match = self.stations.resolve(
                address=address,
                latitude=latitude,
                longitude=longitude,
                legacy_station_id=legacy_station_id,
                reviewed_mapping_method=reviewed_method,
            )
            eligible = match.station_id is not None and not match.review_required
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
                    "generation_start": profile["generation_start"],
                    "generation_end": profile["generation_end"],
                    "source_observation_rows": profile["source_observation_rows"],
                    "source_partition_count": profile["source_partition_count"],
                    "model_ready_status": "eligible" if eligible else "quarantined",
                    "model_ready_reason": None if eligible else match.reason,
                }
            )
        result = pd.DataFrame(rows).sort_values(
            ["company", "plant", "energy_source"], kind="stable"
        ).reset_index(drop=True)
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        result.to_csv(destination, index=False, encoding="utf-8-sig")
        return result

    @staticmethod
    def _address_key(address: object) -> str:
        if address is None or pd.isna(address):
            return ""
        return re.sub(r"[^0-9가-힣]", "", str(address))

    def _scan(self, paths: Iterable[Path]) -> list[dict[str, object]]:
        profiles: dict[tuple[str, str, str], dict[str, object]] = {}
        for path in paths:
            source = pd.read_csv(Path(path), usecols=self.scan_columns)
            source["timestamp"] = pd.to_datetime(source["timestamp"], errors="coerce")
            for key, group in source.groupby(["company", "plant", "energy_source"], sort=False):
                company, plant, energy_source = map(str, key)
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
        result = []
        for profile in profiles.values():
            profile["generation_start"] = profile["generation_start"].isoformat()
            profile["generation_end"] = profile["generation_end"].isoformat()
            profile["source_partition_count"] = len(profile.pop("source_partitions"))
            result.append(profile)
        return result

    @staticmethod
    def _legacy_map(frame: pd.DataFrame | None) -> dict[tuple[str, str], int]:
        if frame is None or frame.empty:
            return {}
        required = {"company", "plant", "station_id"}
        if not required.issubset(frame.columns):
            return {}
        mapping = frame[list(required)].dropna().drop_duplicates()
        ambiguous = mapping.groupby(["company", "plant"]).size()
        ambiguous = ambiguous[ambiguous.gt(1)]
        if not ambiguous.empty:
            raise ValueError(f"Legacy source has ambiguous plant/station mappings: {ambiguous.to_dict()}")
        return {
            (str(row.company), str(row.plant)): int(row.station_id)
            for row in mapping.itertuples(index=False)
        }
