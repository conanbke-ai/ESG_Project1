"""Conversion and validation for the public province boundary layer.

The dashboard consumes a small WGS84 GeoJSON file, while the authoritative
SGIS distribution is a detailed projected Shapefile.  Conversion is therefore
an explicit build step: source metadata is retained, geometries are simplified
in metres, and representative island points are checked before publication.
"""

from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
import json
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence
from zipfile import BadZipFile, ZipFile

from solar_forecast.artifacts.manifest import replace_file_atomic


SGIS_TO_ISO_REGION: Mapping[str, str] = {
    "11": "KR11",  # 서울
    "21": "KR26",  # 부산
    "22": "KR27",  # 대구
    "23": "KR28",  # 인천
    "24": "KR29",  # 광주
    "25": "KR30",  # 대전
    "26": "KR31",  # 울산
    "29": "KR50",  # 세종
    "31": "KR41",  # 경기
    "32": "KR42",  # 강원
    "33": "KR43",  # 충북
    "34": "KR44",  # 충남
    "35": "KR45",  # 전북
    "36": "KR46",  # 전남
    "37": "KR47",  # 경북
    "38": "KR48",  # 경남
    "39": "KR49",  # 제주
}

EXPECTED_REGION_IDS = frozenset(SGIS_TO_ISO_REGION.values())
REQUIRED_COMPONENT_SUFFIXES = (".shp", ".shx", ".dbf", ".prj")


class ProvinceBoundaryError(ValueError):
    """Raised when a province boundary source or output is unsafe to publish."""


@dataclass(frozen=True)
class BoundaryLandmark:
    """A stable interior point used to guard administrative ownership."""

    name: str
    longitude: float
    latitude: float
    expected_region_id: str


BOUNDARY_LANDMARKS: tuple[BoundaryLandmark, ...] = (
    BoundaryLandmark("백령도", 124.7124, 37.9740, "KR28"),
    BoundaryLandmark("대청도", 124.7000, 37.8200, "KR28"),
    BoundaryLandmark("연평도", 125.7000, 37.6600, "KR28"),
    BoundaryLandmark("울릉도", 130.8986, 37.4813, "KR47"),
)


@dataclass(frozen=True)
class SgisBoundarySource:
    """Auditable metadata embedded in a converted SGIS boundary artifact."""

    provider: str
    source_url: str
    reference_date: str
    archive_sha256: str
    shapefile_sha256: str
    dataset_id: str = "sgis_province_boundaries_2025_2q"

    def __post_init__(self) -> None:
        for label, digest in (
            ("archive_sha256", self.archive_sha256),
            ("shapefile_sha256", self.shapefile_sha256),
        ):
            normalized = digest.strip().lower()
            if len(normalized) != 64 or any(
                character not in "0123456789abcdef" for character in normalized
            ):
                raise ProvinceBoundaryError(f"{label} must be a SHA-256 digest")


class SgisProvinceBoundaryConverter:
    """Convert the official SGIS province Shapefile to compact WGS84 GeoJSON."""

    REQUIRED_FIELDS = frozenset({"BASE_DATE", "SIDO_CD", "SIDO_NM"})

    def __init__(self, *, simplify_meters: float = 150.0, precision: int = 6):
        if simplify_meters < 0:
            raise ProvinceBoundaryError("simplify_meters must be non-negative")
        if not 4 <= precision <= 8:
            raise ProvinceBoundaryError("precision must be between 4 and 8")
        self.simplify_meters = float(simplify_meters)
        self.precision = int(precision)

    def convert(
        self,
        source_path: str | Path,
        output_path: str | Path,
        *,
        source_archive_path: str | Path,
        source: SgisBoundarySource,
    ) -> dict[str, Any]:
        """Convert, validate and atomically write one province FeatureCollection."""

        source_path = Path(source_path)
        component_hashes = verify_sgis_source_bundle(
            source_path,
            source_archive_path,
            source=source,
        )
        shapefile, pyproj, shapely_geometry, shapely_ops, shapely_validation = (
            self._geo_dependencies()
        )
        output_path = Path(output_path)
        projection_path = source_path.with_suffix(".prj")
        source_crs = pyproj.CRS.from_wkt(
            projection_path.read_text(encoding="utf-8")
        )
        transformer = pyproj.Transformer.from_crs(
            source_crs, "EPSG:4326", always_xy=True
        )

        reader = shapefile.Reader(str(source_path), encoding="utf-8")
        field_names = {str(field[0]) for field in reader.fields[1:]}
        missing = sorted(self.REQUIRED_FIELDS - field_names)
        if missing:
            raise ProvinceBoundaryError(
                "SGIS province Shapefile is missing fields: " + ", ".join(missing)
            )

        features: list[dict[str, Any]] = []
        seen_codes: set[str] = set()
        base_dates: set[str] = set()
        try:
            for shape_record in reader.iterShapeRecords():
                record = shape_record.record.as_dict()
                base_dates.add(str(record["BASE_DATE"]).strip())
                source_code = str(record["SIDO_CD"]).strip()
                if source_code in seen_codes:
                    raise ProvinceBoundaryError(
                        f"duplicate SGIS province code: {source_code}"
                    )
                seen_codes.add(source_code)
                region_id = SGIS_TO_ISO_REGION.get(source_code)
                if region_id is None:
                    raise ProvinceBoundaryError(
                        f"unknown SGIS province code: {source_code}"
                    )

                geometry = shapely_geometry.shape(
                    shape_record.shape.__geo_interface__
                )
                if not geometry.is_valid:
                    geometry = shapely_validation.make_valid(geometry)
                geometry = _polygonal_geometry(geometry, shapely_ops)
                if self.simplify_meters:
                    geometry = geometry.simplify(
                        self.simplify_meters, preserve_topology=True
                    )
                geometry = shapely_ops.transform(transformer.transform, geometry)
                if geometry.is_empty or geometry.geom_type not in {
                    "Polygon",
                    "MultiPolygon",
                }:
                    raise ProvinceBoundaryError(
                        f"invalid converted geometry for SGIS province {source_code}"
                    )

                features.append(
                    {
                        "type": "Feature",
                        "properties": {
                            "id": region_id,
                            "name": str(record["SIDO_NM"]).strip(),
                            "source_code": source_code,
                            "base_date": str(record["BASE_DATE"]).strip(),
                        },
                        "geometry": _rounded_mapping(
                            shapely_geometry.mapping(geometry), self.precision
                        ),
                    }
                )
        finally:
            reader.close()

        if seen_codes != set(SGIS_TO_ISO_REGION):
            missing_codes = sorted(set(SGIS_TO_ISO_REGION) - seen_codes)
            raise ProvinceBoundaryError(
                "SGIS province Shapefile is incomplete: " + ", ".join(missing_codes)
            )
        expected_base_date = "".join(
            character for character in source.reference_date if character.isdigit()
        )
        if len(expected_base_date) != 8 or base_dates != {expected_base_date}:
            raise ProvinceBoundaryError(
                "SGIS BASE_DATE differs from the declared reference date: "
                f"declared={source.reference_date}, actual={sorted(base_dates)}"
            )

        payload: dict[str, Any] = {
            "type": "FeatureCollection",
            "metadata": {
                "dataset_id": source.dataset_id,
                "provider": source.provider,
                "source_url": source.source_url,
                "reference_date": source.reference_date,
                "archive_sha256": source.archive_sha256.lower(),
                "shapefile_sha256": source.shapefile_sha256.lower(),
                "component_sha256": component_hashes,
                "source_crs": source_crs.to_string(),
                "output_crs": "EPSG:4326",
                "simplify_meters": self.simplify_meters,
                "coordinate_precision": self.precision,
            },
            "features": sorted(
                features, key=lambda feature: feature["properties"]["id"]
            ),
        }
        validate_province_boundaries(payload, require_complete=True)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = output_path.with_name(output_path.name + ".part")
        temporary.write_text(
            json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + "\n",
            encoding="utf-8",
        )
        replace_file_atomic(temporary, output_path)
        return payload

    @staticmethod
    def _geo_dependencies() -> tuple[Any, Any, Any, Any, Any]:
        try:
            import shapefile
            import pyproj
            from shapely import geometry as shapely_geometry
            from shapely import ops as shapely_ops
            from shapely import validation as shapely_validation
        except ImportError as error:
            raise ProvinceBoundaryError(
                "SGIS conversion requires the optional geo dependencies; "
                "install the project with .[geo]"
            ) from error
        return (
            shapefile,
            pyproj,
            shapely_geometry,
            shapely_ops,
            shapely_validation,
        )


def verify_sgis_source_bundle(
    source_path: str | Path,
    source_archive_path: str | Path,
    *,
    source: SgisBoundarySource,
) -> dict[str, str]:
    """Verify the official archive and every Shapefile component used."""

    source_path = Path(source_path)
    archive_path = Path(source_archive_path)
    if not source_path.is_file():
        raise FileNotFoundError(source_path)
    if not archive_path.is_file():
        raise FileNotFoundError(archive_path)
    archive_digest = _file_sha256(archive_path)
    if archive_digest != source.archive_sha256.lower():
        raise ProvinceBoundaryError(
            "SGIS source archive SHA-256 mismatch: "
            f"expected={source.archive_sha256.lower()}, actual={archive_digest}"
        )

    component_paths = {
        suffix.removeprefix("."): source_path.with_suffix(suffix)
        for suffix in REQUIRED_COMPONENT_SUFFIXES
    }
    missing = [str(path) for path in component_paths.values() if not path.is_file()]
    if missing:
        raise ProvinceBoundaryError(
            "SGIS Shapefile components are missing: " + ", ".join(missing)
        )
    component_hashes = {
        name: _file_sha256(path) for name, path in component_paths.items()
    }
    if component_hashes["shp"] != source.shapefile_sha256.lower():
        raise ProvinceBoundaryError(
            "SGIS province Shapefile SHA-256 mismatch: "
            f"expected={source.shapefile_sha256.lower()}, "
            f"actual={component_hashes['shp']}"
        )

    try:
        with ZipFile(archive_path) as archive:
            members: dict[str, list[Any]] = {}
            for info in archive.infolist():
                basename = Path(info.filename).name.casefold()
                members.setdefault(basename, []).append(info)
            for name, path in component_paths.items():
                candidates = members.get(path.name.casefold(), [])
                if not candidates:
                    raise ProvinceBoundaryError(
                        f"SGIS archive does not contain the used {name} component: {path.name}"
                    )
                if not any(
                    _zip_member_sha256(archive, candidate) == component_hashes[name]
                    for candidate in candidates
                ):
                    raise ProvinceBoundaryError(
                        f"SGIS extracted {name} component differs from the verified archive"
                    )
    except BadZipFile as error:
        raise ProvinceBoundaryError("SGIS source archive is not a valid ZIP file") from error
    return component_hashes


def validate_province_boundaries(
    payload: Mapping[str, Any],
    *,
    require_complete: bool = True,
) -> None:
    """Validate GeoJSON structure, unique IDs and representative island owners."""

    if payload.get("type") != "FeatureCollection":
        raise ProvinceBoundaryError("province boundaries must be a FeatureCollection")
    features = payload.get("features")
    if not isinstance(features, list):
        raise ProvinceBoundaryError("province boundary features must be a list")

    by_id: dict[str, Mapping[str, Any]] = {}
    for index, feature in enumerate(features):
        if not isinstance(feature, Mapping) or feature.get("type") != "Feature":
            raise ProvinceBoundaryError(f"invalid province feature at index {index}")
        properties = feature.get("properties")
        geometry = feature.get("geometry")
        if not isinstance(properties, Mapping) or not isinstance(geometry, Mapping):
            raise ProvinceBoundaryError(f"incomplete province feature at index {index}")
        region_id = str(properties.get("id", "")).strip()
        if not region_id:
            raise ProvinceBoundaryError(f"province feature has no id at index {index}")
        if region_id in by_id:
            raise ProvinceBoundaryError(f"duplicate province feature id: {region_id}")
        if geometry.get("type") not in {"Polygon", "MultiPolygon"}:
            raise ProvinceBoundaryError(
                f"unsupported province geometry type for {region_id}"
            )
        if not isinstance(geometry.get("coordinates"), list):
            raise ProvinceBoundaryError(f"missing coordinates for {region_id}")
        by_id[region_id] = feature

    if not require_complete:
        return
    actual_ids = set(by_id)
    if actual_ids != EXPECTED_REGION_IDS:
        missing = sorted(EXPECTED_REGION_IDS - actual_ids)
        extra = sorted(actual_ids - EXPECTED_REGION_IDS)
        raise ProvinceBoundaryError(
            f"province boundary IDs differ; missing={missing}, extra={extra}"
        )

    for landmark in BOUNDARY_LANDMARKS:
        owners = [
            region_id
            for region_id, feature in by_id.items()
            if geometry_contains_point(
                feature["geometry"], landmark.longitude, landmark.latitude
            )
        ]
        if owners != [landmark.expected_region_id]:
            raise ProvinceBoundaryError(
                f"{landmark.name} must belong only to {landmark.expected_region_id}; "
                f"actual={owners}"
            )


def geometry_contains_point(
    geometry: Mapping[str, Any], longitude: float, latitude: float
) -> bool:
    """Return whether a WGS84 Polygon/MultiPolygon contains a point."""

    coordinates = geometry.get("coordinates", [])
    geometry_type = geometry.get("type")
    polygons = [coordinates] if geometry_type == "Polygon" else coordinates
    if geometry_type not in {"Polygon", "MultiPolygon"}:
        return False
    return any(
        _polygon_contains_point(polygon, longitude, latitude)
        for polygon in polygons
    )


def _polygon_contains_point(
    polygon: Sequence[Any], longitude: float, latitude: float
) -> bool:
    if not polygon or not _ring_contains_point(polygon[0], longitude, latitude):
        return False
    return not any(
        _ring_contains_point(hole, longitude, latitude) for hole in polygon[1:]
    )


def _ring_contains_point(
    ring: Sequence[Sequence[float]], longitude: float, latitude: float
) -> bool:
    if len(ring) < 4:
        return False
    inside = False
    previous = ring[-1]
    for current in ring:
        x1, y1 = float(previous[0]), float(previous[1])
        x2, y2 = float(current[0]), float(current[1])
        crosses = (y1 > latitude) != (y2 > latitude)
        if crosses:
            intersection = (x2 - x1) * (latitude - y1) / (y2 - y1) + x1
            if longitude < intersection:
                inside = not inside
        previous = current
    return inside


def _polygonal_geometry(geometry: Any, shapely_ops: Any) -> Any:
    if geometry.geom_type in {"Polygon", "MultiPolygon"}:
        return geometry
    if geometry.geom_type == "GeometryCollection":
        polygons = [
            item
            for item in geometry.geoms
            if item.geom_type in {"Polygon", "MultiPolygon"}
        ]
        if polygons:
            return shapely_ops.unary_union(polygons)
    raise ProvinceBoundaryError(
        f"SGIS geometry is not polygonal after repair: {geometry.geom_type}"
    )


def _rounded_mapping(value: Any, precision: int) -> Any:
    if isinstance(value, Mapping):
        return {key: _rounded_mapping(item, precision) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_rounded_mapping(item, precision) for item in value]
    if isinstance(value, float):
        return round(value, precision)
    return value


def _file_sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as source:
        while chunk := source.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def _zip_member_sha256(archive: ZipFile, member: Any) -> str:
    digest = sha256()
    with archive.open(member) as source:
        while chunk := source.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


__all__ = [
    "BOUNDARY_LANDMARKS",
    "EXPECTED_REGION_IDS",
    "BoundaryLandmark",
    "ProvinceBoundaryError",
    "SgisBoundarySource",
    "SgisProvinceBoundaryConverter",
    "geometry_contains_point",
    "validate_province_boundaries",
    "verify_sgis_source_bundle",
]
