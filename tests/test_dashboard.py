import csv
from hashlib import sha256
import json
from pathlib import Path

import pandas as pd
import pytest

from solar_forecast.cli import build_parser
from solar_forecast.collectors.csv_artifacts import write_standardized_csv
from solar_forecast.reporting import DashboardBuilder, ModelAnalyticsService
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


def _write_static_dashboard(root: Path) -> dict[str, str]:
    static_root = root / "dashboard"
    (static_root / "assets").mkdir(parents=True, exist_ok=True)
    static_files = {
        "solar_dashboard.html": "<html>national</html>",
        "model_analysis.html": "<html>analytics</html>",
        "plant_region_report_perm.html": "<html>redirect</html>",
        "assets/dashboard.css": "body {}",
        "assets/dashboard.js": "void 0;",
    }
    for relative_path, content in static_files.items():
        (static_root / relative_path).write_text(content, encoding="utf-8")
    return static_files


def _write_model_run(
    root: Path,
    *,
    model: str,
    run: str,
    contract: dict,
    test_frame: pd.DataFrame,
    calibration_frame: pd.DataFrame,
    mode: str = "full",
    status: str = "completed",
    energy_source: str = "solar",
) -> None:
    path = root / "artifacts/models" / model / run / "manifest.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    prediction_column = "xgb_pred" if model == "xgboost" else "cnn_pred"
    test_predictions = path.parent / "test_predictions.csv"
    calibration_predictions = path.parent / "calibration_predictions.csv"
    test_frame.rename(columns={"y_pred": prediction_column}).to_csv(
        test_predictions, index=False, encoding="utf-8-sig"
    )
    calibration_frame.rename(columns={"y_pred": prediction_column}).to_csv(
        calibration_predictions, index=False, encoding="utf-8-sig"
    )
    path.write_text(
        json.dumps(
            {
                "status": status,
                "model": model,
                "details": {
                    "run_context": {
                        "execution_mode": mode,
                        "energy_source": energy_source,
                        "target": "generation_mwh",
                        "target_unit": "MWh",
                    },
                    "metrics": {"mae": 0.2, "rmse": 0.3, "r2": 0.8},
                    "n_test": len(test_frame),
                    "evaluation_contract": contract,
                    "test_predictions": str(test_predictions),
                    "calibration_predictions": str(calibration_predictions),
                },
            }
        ),
        encoding="utf-8",
    )


def test_serve_dashboard_defaults_to_expected_local_url():
    args = build_parser().parse_args(["serve-dashboard"])

    assert args.host == "127.0.0.1"
    assert args.port == 5500
    assert args.output_dir == "dashboard"
    assert args.no_refresh is False


def test_dashboard_builder_publishes_clean_user_contract_without_registry(tmp_path):
    _write_national_inventory_fixture(tmp_path)
    static_files = _write_static_dashboard(tmp_path)
    quality = pd.DataFrame(
        [
            {
                "plant_id": "solar:review",
                "plant": "검토 태양광",
                "region": "전라남도",
                "energy_source": "solar",
                "rows": 1000,
                "start": "2024-01-01T00:00:00",
                "end": "2024-12-31T23:00:00",
                "sensor_risk": "review",
                "missing_weather_rate": 0.4,
                "daylight_zero_rate": 0.1,
                "capacity_exceeded_rate": 0.0,
                "positive_flatline_rate": 0.0,
                "temporal_profile_consistency": 0.9,
                "peer_pattern_correlation": 0.8,
            },
            {
                "plant_id": "wind:review",
                "plant": "검토 풍력",
                "region": "전라남도",
                "energy_source": "wind",
                "rows": 1000,
                "sensor_risk": "review",
            },
        ]
    )
    write_standardized_csv(
        quality, tmp_path / "file/standardized/plant_quality_report.csv"
    )

    result = DashboardBuilder(tmp_path, tmp_path / "published").build()
    payload = json.loads(result.data_path.read_text(encoding="utf-8"))

    assert set(payload) == {"meta", "national_inventory", "model_analysis"}
    assert not {
        "mapping",
        "data_inventory",
        "feature_contract",
        "plants",
    }.intersection(payload)
    assert set(payload["national_inventory"]["source"]) == {
        "provider",
        "source_system",
        "source_url",
        "reference_date",
        "scope",
    }
    assert payload["national_inventory"]["summary"]["generator_records"] == 1
    region = next(
        row
        for row in payload["national_inventory"]["regions"]
        if row["region"] == "전라남도"
    )
    assert region["local_area_count"] == 1
    assert payload["model_analysis"]["status"] == "empty"
    assert [
        signal["plant"]
        for signal in payload["model_analysis"]["anomalies"][
            "data_quality_signals"
        ]
    ] == ["검토 태양광"]
    assert payload["model_analysis"]["anomalies"]["data_quality_signals"][0][
        "plant_id"
    ] == "solar:review"
    assert result.national_generator_records == 1
    assert result.national_capacity_mw == 1.25
    assert result.model_analysis_status == "empty"
    assert result.data_quality_signals == 1
    assert result.analytics_dashboard == tmp_path / "published/model_analysis.html"
    assert result.mapping_report == result.analytics_dashboard
    assert result.boundary_path == tmp_path / "published/data/korea_provinces.geojson"
    for relative_path, content in static_files.items():
        assert (tmp_path / "published" / relative_path).read_text(
            encoding="utf-8"
        ) == content


def test_model_analytics_excludes_smoke_and_requires_matching_contract(tmp_path):
    quality_path = tmp_path / "file/standardized/plant_quality_report.csv"
    write_standardized_csv(
        pd.DataFrame(columns=["plant", "region", "energy_source", "sensor_risk"]),
        quality_path,
    )
    contract = {
        "dataset_fingerprint": "same-dataset",
        "target": "generation_mwh",
        "target_unit": "MWh",
        "horizon_hours": 24,
        "test_start": "2025-01-01T00:00:00",
        "test_end": "2025-02-01T00:00:00",
        "prediction_key": ["timestamp", "plant_id"],
    }

    write_standardized_csv(
        pd.DataFrame(
            {
                "plant_id": ["p1", "p2"],
                "capacity_mw": [2.0, None],
            }
        ),
        tmp_path / "file/standardized/plant_registry.csv",
    )

    def manifest(
        model: str,
        run: str,
        *,
        mode: str,
        status: str = "completed",
        energy_source: str = "solar",
    ):
        path = tmp_path / "artifacts/models" / model / run / "manifest.json"
        path.parent.mkdir(parents=True, exist_ok=True)
        prediction_column = "xgb_pred" if model == "xgboost" else "cnn_pred"
        test_predictions = path.parent / "test_predictions.csv"
        calibration_predictions = path.parent / "calibration_predictions.csv"
        pd.DataFrame(
            {
                "timestamp": list(
                    pd.date_range("2025-01-01", periods=2, freq="h").repeat(2)
                ),
                "plant_id": ["p1", "p2", "p1", "p2"],
                "region": ["전라남도", "경기도", "전라남도", "경기도"],
                "plant": ["영암", "화성", "영암", "화성"],
                "split": ["test"] * 4,
                "y_true": [1.0, 2.0, 1.5, 2.5],
                prediction_column: (
                    [0.9, 2.1, 1.4, 2.4]
                    if model == "xgboost"
                    else [1.1, 1.8, 1.7, 2.6]
                ),
            }
        ).to_csv(test_predictions, index=False, encoding="utf-8-sig")
        pd.DataFrame(
            {
                "plant_id": ["p1", "p2", "p1", "p2", "p1", "p2"],
                "y_true": [1.0] * 6,
                prediction_column: [0.99, 1.01, 0.98, 1.02, 0.97, 1.03],
            }
        ).to_csv(calibration_predictions, index=False, encoding="utf-8-sig")
        path.write_text(
            json.dumps(
                {
                    "status": status,
                    "model": model,
                    "details": {
                        "run_context": {
                            "execution_mode": mode,
                            "energy_source": energy_source,
                            "target": "generation_mwh",
                            "target_unit": "MWh",
                        },
                        "metrics": {"mae": 0.2, "rmse": 0.3, "r2": 0.8},
                        "n_test": 4,
                        "evaluation_contract": contract,
                        "test_predictions": str(test_predictions),
                        "calibration_predictions": str(calibration_predictions),
                    },
                }
            ),
            encoding="utf-8",
        )

    manifest("xgboost", "20260101_000000", mode="smoke")
    manifest("xgboost", "20260102_000000", mode="full")
    manifest("xgboost", "20260104_000000", mode="full", energy_source="wind")
    manifest("cnn_bilstm", "20260102_000000", mode="full")
    manifest("cnn_bilstm", "20260103_000000", mode="full", status="failed")

    analytics = ModelAnalyticsService(tmp_path).build()

    assert analytics["status"] == "ready"
    assert analytics["evaluation"] == {
        "scope": "test",
        "from": "2025-01-01T00:00:00",
        "to": "2025-02-01T00:00:00",
        "horizon_hours": 24,
        "common_samples": 4,
    }
    assert {row["id"] for row in analytics["models"]} == {
        "xgboost",
        "cnn_bilstm",
    }
    assert all(row["comparable"] for row in analytics["models"])
    assert len(analytics["regions"]) == 4
    assert len(analytics["plants"]) == 4
    assert len(analytics["series"]) == 4
    for model in analytics["models"]:
        assert model["metrics"]["nmae_capacity"] is not None
        assert model["metrics"]["capacity_samples"] == 2
        assert model["metrics"]["capacity_coverage"] == 0.5
    metrics_by_model = {
        model["id"]: model["metrics"] for model in analytics["models"]
    }
    assert metrics_by_model["xgboost"]["nmae_capacity"] == pytest.approx(5.0)
    assert metrics_by_model["cnn_bilstm"]["nmae_capacity"] == pytest.approx(7.5)
    assert analytics["anomalies"]["prediction_summary"][
        "evaluated_predictions"
    ] == 4


def test_model_analytics_uses_plant_calibration_and_unit_safe_event_ranking(
    tmp_path,
):
    contract = {
        "dataset_fingerprint": "calibrated-dataset",
        "target": "generation_mwh",
        "target_unit": "MWh",
        "horizon_hours": 24,
        "test_start": "2025-01-01T00:00:00",
        "test_end": "2025-01-02T00:00:00",
        "prediction_key": ["timestamp", "plant_id"],
    }
    write_standardized_csv(
        pd.DataFrame(
            {
                "plant_id": ["known", "missing"],
                "capacity_mw": [100.0, None],
            }
        ),
        tmp_path / "file/standardized/plant_registry.csv",
    )
    test_frame = pd.DataFrame(
        {
            "timestamp": pd.date_range("2025-01-01", periods=2, freq="h"),
            "plant_id": ["known", "missing"],
            "region": ["전라남도", "전라남도"],
            "plant": ["용량확인", "용량미확인"],
            "split": ["test", "test"],
            "y_true": [0.0, 0.0],
            "y_pred": [1.0, 5.0],
        }
    )
    calibration_frame = pd.DataFrame(
        {
            "plant_id": ["known"] * 168 + ["missing"] * 168,
            "split": ["calibration"] * 336,
            "y_true": [0.0] * 336,
            "y_pred": [0.1] * 168 + [1.0] * 168,
        }
    )
    for model in ("xgboost", "cnn_bilstm"):
        _write_model_run(
            tmp_path,
            model=model,
            run="20260101_000000",
            contract=contract,
            test_frame=test_frame,
            calibration_frame=calibration_frame,
        )

    analytics = ModelAnalyticsService(tmp_path).build()

    events = analytics["anomalies"]["prediction_signals"]
    assert analytics["status"] == "ready"
    assert len(events) == 4
    assert [event["plant_id"] for event in events[:2]] == ["known", "known"]
    assert events[0]["absolute_error"] < events[2]["absolute_error"]
    assert events[0]["exceedance_ratio"] > events[2]["exceedance_ratio"]
    assert {event["threshold_source"] for event in events} == {
        "plant_capacity_normalized",
        "plant_absolute_capacity_missing",
    }
    summary = analytics["anomalies"]["prediction_summary"]
    assert summary["total"] == 4
    assert summary["evaluated_predictions"] == 4
    assert {row["signals"] for row in summary["by_model"]} == {2}


def test_model_analytics_falls_back_to_global_capacity_normalized_threshold(
    tmp_path,
):
    contract = {
        "dataset_fingerprint": "global-normalized-fallback",
        "target": "generation_mwh",
        "target_unit": "MWh",
        "horizon_hours": 24,
        "test_start": "2025-01-01T00:00:00",
        "test_end": "2025-01-01T00:00:00",
        "prediction_key": ["timestamp", "plant_id"],
    }
    write_standardized_csv(
        pd.DataFrame(
            {
                "plant_id": ["small", "large"],
                "capacity_mw": [10.0, 100.0],
            }
        ),
        tmp_path / "file/standardized/plant_registry.csv",
    )
    test_frame = pd.DataFrame(
        {
            "timestamp": [pd.Timestamp("2025-01-01")] * 2,
            "plant_id": ["small", "large"],
            "region": ["전라남도", "전라남도"],
            "plant": ["소형", "대형"],
            "split": ["test", "test"],
            "y_true": [0.0, 0.0],
            "y_pred": [0.2, 2.0],
        }
    )
    calibration_frame = pd.DataFrame(
        {
            "plant_id": ["small"] * 100 + ["large"] * 100,
            "split": ["calibration"] * 200,
            "y_true": [0.0] * 200,
            "y_pred": [0.1] * 100 + [1.0] * 100,
        }
    )
    for model in ("xgboost", "cnn_bilstm"):
        _write_model_run(
            tmp_path,
            model=model,
            run="20260101_000000",
            contract=contract,
            test_frame=test_frame,
            calibration_frame=calibration_frame,
        )

    analytics = ModelAnalyticsService(tmp_path).build()

    events = analytics["anomalies"]["prediction_signals"]
    assert analytics["status"] == "ready"
    assert len(events) == 4
    assert {event["threshold_source"] for event in events} == {
        "global_capacity_normalized"
    }


def test_model_analytics_caps_events_but_counts_all_and_keeps_latest_hourly_segment(
    tmp_path,
):
    contract = {
        "dataset_fingerprint": "large-evaluation",
        "target": "generation_mwh",
        "target_unit": "MWh",
        "horizon_hours": 24,
        "test_start": "2025-01-01T00:00:00",
        "test_end": "2025-02-01T00:00:00",
        "prediction_key": ["timestamp", "plant_id"],
    }
    write_standardized_csv(
        pd.DataFrame({"plant_id": ["p1"], "capacity_mw": [10.0]}),
        tmp_path / "file/standardized/plant_registry.csv",
    )
    first_segment = pd.date_range("2025-01-01", periods=200, freq="h")
    second_segment = pd.date_range(
        first_segment[-1] + pd.Timedelta(hours=2), periods=200, freq="h"
    )
    timestamps = first_segment.append(second_segment)
    test_frame = pd.DataFrame(
        {
            "timestamp": timestamps,
            "plant_id": ["p1"] * 400,
            "region": ["전라남도"] * 400,
            "plant": ["연속성검증"] * 400,
            "split": ["test"] * 400,
            "y_true": [0.0] * 400,
            "y_pred": [1.0] * 400,
        }
    )
    calibration_frame = pd.DataFrame(
        {
            "plant_id": ["p1"] * 168,
            "split": ["calibration"] * 168,
            "y_true": [0.0] * 168,
            "y_pred": [0.1] * 168,
        }
    )
    for model in ("xgboost", "cnn_bilstm"):
        _write_model_run(
            tmp_path,
            model=model,
            run="20260101_000000",
            contract=contract,
            test_frame=test_frame,
            calibration_frame=calibration_frame,
        )

    analytics = ModelAnalyticsService(tmp_path).build()

    events = analytics["anomalies"]["prediction_signals"]
    summary = analytics["anomalies"]["prediction_summary"]
    assert len(events) == 500
    assert summary["total"] == 800
    assert summary["returned_top_events"] == 500
    assert summary["evaluated_predictions"] == 800
    assert [row["signals"] for row in summary["by_model"]] == [400, 400]
    assert summary["by_region"][0]["signals"] == 800
    assert summary["by_plant"][0]["signals"] == 800
    assert len(analytics["series"]) == 168
    assert analytics["series"][0]["timestamp"] == second_segment[32].strftime(
        "%Y-%m-%dT%H:%M:%S"
    )


def test_model_analytics_selects_latest_complete_compatible_signature(tmp_path):
    common = {
        "target": "generation_mwh",
        "target_unit": "MWh",
        "horizon_hours": 24,
        "prediction_key": ["timestamp", "plant_id"],
    }
    compatible_contract = {
        **common,
        "dataset_fingerprint": "compatible",
        "test_start": "2025-01-01T00:00:00",
        "test_end": "2025-01-03T00:00:00",
    }
    newer_unpaired_contract = {
        **common,
        "dataset_fingerprint": "newer-unpaired",
        "test_start": "2025-02-01T00:00:00",
        "test_end": "2025-02-03T00:00:00",
    }
    write_standardized_csv(
        pd.DataFrame({"plant_id": ["p1"], "capacity_mw": [10.0]}),
        tmp_path / "file/standardized/plant_registry.csv",
    )
    test_frame = pd.DataFrame(
        {
            "timestamp": pd.date_range("2025-01-01", periods=3, freq="h"),
            "plant_id": ["p1"] * 3,
            "region": ["전라남도"] * 3,
            "plant": ["호환발전소"] * 3,
            "split": ["test"] * 3,
            "y_true": [1.0] * 3,
            "y_pred": [0.9] * 3,
        }
    )
    calibration_frame = pd.DataFrame(
        {
            "plant_id": ["p1"] * 168,
            "split": ["calibration"] * 168,
            "y_true": [1.0] * 168,
            "y_pred": [0.99] * 168,
        }
    )
    _write_model_run(
        tmp_path,
        model="xgboost",
        run="20260101_000000",
        contract=compatible_contract,
        test_frame=test_frame,
        calibration_frame=calibration_frame,
    )
    _write_model_run(
        tmp_path,
        model="cnn_bilstm",
        run="20260102_000000",
        contract=compatible_contract,
        test_frame=test_frame,
        calibration_frame=calibration_frame,
    )
    _write_model_run(
        tmp_path,
        model="xgboost",
        run="20260201_000000",
        contract=newer_unpaired_contract,
        test_frame=test_frame,
        calibration_frame=calibration_frame,
    )

    analytics = ModelAnalyticsService(tmp_path).build()

    assert analytics["status"] == "ready"
    assert analytics["evaluation"]["from"] == "2025-01-01T00:00:00"
    assert analytics["evaluation"]["common_samples"] == 3


def test_dashboard_frontend_has_no_retired_developer_quality_view():
    root = Path(__file__).resolve().parents[1]
    script = (root / "dashboard/assets/dashboard.js").read_text(encoding="utf-8")
    analysis = (root / "dashboard/model_analysis.html").read_text(encoding="utf-8")

    for retired in (
        "plant_id_unique",
        "원본 CSV",
        "파일 포맷 판단",
        "인코딩",
        "학습 매핑·품질",
        "drawTrainingMap",
    ):
        assert retired not in script
        assert retired not in analysis
    assert "model_analysis" in script
    assert "source_region_conflict" in script
    assert ".slice(0, 3)" not in script


def test_dashboard_frontend_compares_all_metrics_and_supports_national_search():
    root = Path(__file__).resolve().parents[1]
    script = (root / "dashboard/assets/dashboard.js").read_text(encoding="utf-8")
    styles = (root / "dashboard/assets/dashboard.css").read_text(encoding="utf-8")
    coverage = (root / "dashboard/solar_dashboard.html").read_text(encoding="utf-8")

    assert "analysis-metric" not in script
    assert "state.metric" not in script
    assert "metric-overview-grid" in script
    assert "metricQuartet" in script
    for metric in ("NMAE", "MAE", "RMSE", "R²"):
        assert metric in script

    assert "L.tileLayer" in script
    assert 'L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png"' in script
    assert "L.control.zoom" in script
    assert "zoomSnap: .1" in script
    assert "openstreetmap.org" in script
    assert "tile.openstreetmap.org" in coverage
    assert "province-hover-tooltip" in script
    assert "permanent: true" not in script

    assert "전국 세부지역 검색" in script
    assert "shortRegion(row.region)" in script
    assert "regionSearchTerms" in script
    assert "전라북도" in script
    assert "data-detail-region" in script
    assert "<th>시도</th>" in script
    assert "summary-value" in script
    assert ".summary-value" in styles
    assert "comparable.length < 2" in script
    assert 'role="region" aria-label="전국 시도별 태양광 설비 분포 지도"' in script
