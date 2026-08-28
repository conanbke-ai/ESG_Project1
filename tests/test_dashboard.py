import csv
from hashlib import sha256
import json
from pathlib import Path

import pandas as pd

from solar_forecast.cli import build_parser
from solar_forecast.collectors.csv_artifacts import write_standardized_csv
from solar_forecast.reporting import DashboardBuilder
from solar_forecast.reporting.national_inventory import REQUIRED_COLUMNS


def _write_national_inventory_fixture(root: Path) -> None:
    source = root / "file/generator_file/kpx_solar_20260828.csv"
    source.parent.mkdir(parents=True, exist_ok=True)
    with source.open("w", encoding="cp949", newline="") as target:
        writer = csv.writer(target)
        writer.writerow(REQUIRED_COLUMNS)
        writer.writerow(
            [
                "사업자",
                "발전기",
                "1",
                "1.25",
                "비회원",
                "비중앙",
                "태양에너지",
                "신재생",
                "발전사업",
                "전남",
                "전라남도 영암군",
            ]
        )
        writer.writerow(["", "통합", "1", "1.25", "", "", "", "", "", "", ""])
    boundary = root / "map/json/custom_boundaries.json"
    boundary.parent.mkdir(parents=True, exist_ok=True)
    boundary.write_text(
        json.dumps({"type": "FeatureCollection", "features": []}),
        encoding="utf-8",
    )
    config = root / "config/national_solar_inventory.json"
    config.parent.mkdir(parents=True, exist_ok=True)
    config.write_text(
        json.dumps(
            {
                "provider": "한국전력거래소",
                "source_system": "EPSIS",
                "source_url": "https://example.test/epsis",
                "local_path": str(source.relative_to(root)).replace("\\", "/"),
                "reference_date": "2026-08-05",
                "downloaded_at": "2026-08-28T10:00:00+09:00",
                "expected_sha256": sha256(source.read_bytes()).hexdigest(),
                "encoding": "cp949",
                "scope": "전국 태양광 발전기 등록 레코드",
                "limitations": ["테스트 원천"],
                "boundary_path": str(boundary.relative_to(root)).replace("\\", "/"),
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def test_serve_dashboard_defaults_to_expected_local_url():
    args = build_parser().parse_args(["serve-dashboard"])

    assert args.host == "127.0.0.1"
    assert args.port == 5500
    assert args.output_dir == "dashboard"
    assert args.no_refresh is False


def test_national_dashboard_keeps_audit_details_offscreen_and_uses_choropleth():
    script = (
        Path(__file__).resolve().parents[1] / "dashboard/assets/dashboard.js"
    ).read_text(encoding="utf-8")
    coverage_view = script.split("function coverageView", 1)[1].split(
        "function qualityView", 1
    )[0]
    national_map = script.split("function drawNationalMap", 1)[1].split(
        "function drawTrainingMap", 1
    )[0]

    assert "수집 범위와 출처" not in coverage_view
    assert "원천 품질 점검" not in coverage_view
    assert 'data-map-metric="capacity"' in coverage_view
    assert "L.tileLayer" not in national_map
    assert "L.circleMarker" not in national_map
    assert "L.geoJSON" in national_map


def test_dashboard_builder_separates_plant_and_weather_proxy_coordinates(tmp_path):
    _write_national_inventory_fixture(tmp_path)
    standardized = tmp_path / "file/standardized"
    weather = tmp_path / "file/KMA_data_file"
    archive = tmp_path / "file/solar_data_file/한국남부발전"
    registry = pd.DataFrame(
        [
            {
                "plant_id": "kospo:exact",
                "company": "kospo",
                "plant": "실좌표 발전소",
                "energy_source": "solar",
                "capacity_mw": 1.0,
                "admin_province": "전라남도",
                "admin_city": "영암군",
                "latitude": 34.8,
                "longitude": 126.7,
                "weather_station_id": 165,
                "weather_station_name": "목포",
                "weather_mapping_method": "coordinate_nearest",
                "weather_mapping_confidence": "high",
                "generation_start": "2024-01-01T00:00:00",
                "generation_end": "2024-12-31T23:00:00",
                "source_observation_rows": 8784,
                "model_ready_status": "eligible",
                "model_ready_reason": None,
            },
            {
                "plant_id": "kospo:proxy",
                "company": "kospo",
                "plant": "대리좌표 발전소",
                "energy_source": "solar",
                "capacity_mw": 2.0,
                "admin_province": "전라남도",
                "admin_city": "영암군",
                "latitude": None,
                "longitude": None,
                "weather_station_id": 165,
                "weather_station_name": "목포",
                "weather_mapping_method": "reviewed_config_mapping",
                "weather_mapping_confidence": "high",
                "generation_start": "2024-01-01T00:00:00",
                "generation_end": "2024-12-31T23:00:00",
                "source_observation_rows": 8784,
                "model_ready_status": "eligible",
                "model_ready_reason": None,
            },
        ]
    )
    quality = pd.DataFrame(
        {
            "plant_id": ["kospo:exact", "kospo:proxy"],
            "rows": [8784, 8784],
            "hourly_coverage": [1.0, 1.0],
            "missing_weather_rate": [0.0, 0.0],
            "sensor_risk": ["low", "low"],
        }
    )
    station = pd.DataFrame(
        {
            "지점": [165],
            "시작일": ["1964-01-01"],
            "지점명": ["목포"],
            "위도": [34.8169],
            "경도": [126.3812],
        }
    )
    write_standardized_csv(registry, standardized / "plant_registry.csv")
    write_standardized_csv(quality, standardized / "plant_quality_report.csv")
    write_standardized_csv(station, weather / "META_관측지점정보.csv")
    standardized.mkdir(parents=True, exist_ok=True)
    (standardized / "model_ready_manifest.json").write_text(
        json.dumps({"rows": 17568, "plants": 2, "features": ["hour"]}),
        encoding="utf-8",
    )
    archive.mkdir(parents=True, exist_ok=True)
    (archive / "한국남부발전(주)_[남제주소내] 태양광발전실적_20250228.csv").write_bytes(
        "발전소\n남제주소내\n".encode("cp949")
    )
    static_root = tmp_path / "dashboard"
    (static_root / "assets").mkdir(parents=True, exist_ok=True)
    static_files = {
        "solar_dashboard.html": "<html>national</html>",
        "plant_region_report_perm.html": "<html>quality</html>",
        "assets/dashboard.css": "body {}",
        "assets/dashboard.js": "void 0;",
    }
    for relative_path, content in static_files.items():
        (static_root / relative_path).write_text(content, encoding="utf-8")

    result = DashboardBuilder(tmp_path, tmp_path / "published").build()
    payload = json.loads(result.data_path.read_text(encoding="utf-8"))

    assert result.solar_assets == 2
    assert result.eligible_solar_assets == 2
    assert result.national_generator_records == 1
    assert result.national_capacity_mw == 1.25
    assert result.boundary_path == tmp_path / "published/data/korea_provinces.geojson"
    assert result.boundary_path.read_text(encoding="utf-8").startswith("{")
    for relative_path, content in static_files.items():
        assert (tmp_path / "published" / relative_path).read_text(
            encoding="utf-8"
        ) == content
    assert payload["national_inventory"]["summary"]["generator_records"] == 1
    assert payload["mapping"]["location_basis_counts"] == {
        "plant_coordinate": 1,
        "weather_station_proxy": 1,
    }
    assert payload["mapping"]["validation"] == [
        {"check": "plant_id_unique", "passed": True, "violations": 0},
        {"check": "eligible_solar_has_weather_station", "passed": True, "violations": 0},
        {"check": "coordinate_pairs_complete", "passed": True, "violations": 0},
        {"check": "coordinates_inside_korea_bounds", "passed": True, "violations": 0},
        {"check": "unresolved_assets_are_quarantined", "passed": True, "violations": 0},
    ]
    assert payload["data_inventory"]["encoding_counts"] == {"cp949": 1}
    assert payload["data_inventory"]["filename_counts"] == {"canonical": 1}
