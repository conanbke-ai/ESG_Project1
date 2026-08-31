from copy import deepcopy
from hashlib import sha256
import json
from pathlib import Path
from zipfile import ZipFile

import pytest

from solar_forecast.cli import build_parser
from solar_forecast.reporting.province_boundaries import (
    BOUNDARY_LANDMARKS,
    ProvinceBoundaryError,
    SgisBoundarySource,
    geometry_contains_point,
    validate_province_boundaries,
    verify_sgis_source_bundle,
)


def _official_boundaries() -> dict:
    root = Path(__file__).resolve().parents[1]
    return json.loads(
        (root / "map/json/geoJson.json").read_text(encoding="utf-8")
    )


def test_official_sgis_boundaries_have_complete_unique_regions_and_island_owners():
    payload = _official_boundaries()

    validate_province_boundaries(payload, require_complete=True)

    assert payload["metadata"]["dataset_id"] == "sgis_province_boundaries_2025_2q"
    assert payload["metadata"]["reference_date"] == "2025-06-30"
    assert payload["metadata"]["output_crs"] == "EPSG:4326"
    for landmark in BOUNDARY_LANDMARKS:
        owners = [
            feature["properties"]["id"]
            for feature in payload["features"]
            if geometry_contains_point(
                feature["geometry"], landmark.longitude, landmark.latitude
            )
        ]
        assert owners == [landmark.expected_region_id], landmark.name


def test_boundary_validation_rejects_duplicate_region_ids():
    payload = deepcopy(_official_boundaries())
    incheon = next(
        feature
        for feature in payload["features"]
        if feature["properties"]["id"] == "KR28"
    )
    incheon["properties"]["id"] = "KR41"

    with pytest.raises(ProvinceBoundaryError, match="duplicate province feature id"):
        validate_province_boundaries(payload, require_complete=True)


def test_prepare_boundaries_cli_requires_auditable_source_digests():
    args = build_parser().parse_args(
        [
            "prepare-boundaries",
            "--source-shp",
            "official.shp",
            "--source-archive",
            "official.zip",
            "--reference-date",
            "2025-06-30",
            "--archive-sha256",
            "a" * 64,
            "--shapefile-sha256",
            "b" * 64,
        ]
    )

    assert args.output == "map/json/geoJson.json"
    assert args.source_archive == "official.zip"
    assert args.simplify_meters == 150.0
    assert args.precision == 6


def test_sgis_source_bundle_verifies_archive_and_every_sidecar(tmp_path: Path):
    source_path = tmp_path / "official.shp"
    component_bytes = {
        ".shp": b"shape",
        ".shx": b"index",
        ".dbf": b"records",
        ".prj": b"projection",
    }
    for suffix, content in component_bytes.items():
        source_path.with_suffix(suffix).write_bytes(content)
    archive_path = tmp_path / "official.zip"
    with ZipFile(archive_path, "w") as archive:
        for suffix, content in component_bytes.items():
            archive.writestr(f"nested/official{suffix}", content)
    source = SgisBoundarySource(
        provider="official",
        source_url="https://example.test/official",
        reference_date="2025-06-30",
        archive_sha256=sha256(archive_path.read_bytes()).hexdigest(),
        shapefile_sha256=sha256(component_bytes[".shp"]).hexdigest(),
    )

    hashes = verify_sgis_source_bundle(
        source_path,
        archive_path,
        source=source,
    )

    assert set(hashes) == {"shp", "shx", "dbf", "prj"}
    source_path.with_suffix(".dbf").write_bytes(b"tampered")
    with pytest.raises(ProvinceBoundaryError, match="differs from the verified archive"):
        verify_sgis_source_bundle(source_path, archive_path, source=source)
