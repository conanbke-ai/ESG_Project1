import json
from pathlib import Path

import pandas as pd

from solar_forecast.collectors.csv_artifacts import write_standardized_csv
from solar_forecast.reporting import DashboardBuilder


def test_dashboard_builder_separates_plant_and_weather_proxy_coordinates(tmp_path):
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

    result = DashboardBuilder(tmp_path, tmp_path / "dashboard").build()
    payload = json.loads(result.data_path.read_text(encoding="utf-8"))

    assert result.solar_assets == 2
    assert result.eligible_solar_assets == 2
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
