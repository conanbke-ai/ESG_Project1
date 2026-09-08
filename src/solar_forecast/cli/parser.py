"""Parser: CLI input translation and command dispatch."""
from __future__ import annotations

from solar_forecast.infrastructure.project_paths import BRONZE_ROOT, COLLECTOR_SILVER_ROOT, FEATURE_EVALUATION_ROOT, GENERATION_ARCHIVE_ROOT, GOLD_DATASET_PATH, HYBRID_EXPERIMENT_ROOT, LEGACY_MERGED_ROOT, LEGACY_MERGED_SOURCE, PIPELINE_RUNS_ROOT, SILVER_ROOT, VERIFICATION_ROOT, WEATHER_ARCHIVE_ROOT

import argparse
from solar_forecast.jobs.contracts import list_job_contracts
from solar_forecast.cli.arguments import add_sequence_arguments
from solar_forecast.cli.dataset_commands import handle_audit_candidate_data_command
from solar_forecast.cli.dashboard_commands import handle_build_dashboard_command
from solar_forecast.cli.collection_commands import handle_collect_command
from solar_forecast.cli.training_commands import (
    handle_evaluate_features_command,
    handle_hybrid_command,
)
from solar_forecast.cli.job_commands import handle_job_contract_command, handle_jobs_command
from solar_forecast.cli.notification_commands import handle_notify_anomalies_command
from solar_forecast.cli.training_commands import handle_pipeline_command
from solar_forecast.cli.dashboard_commands import handle_prepare_boundaries_command
from solar_forecast.cli.dataset_commands import handle_prepare_data_command
from solar_forecast.cli.dashboard_commands import handle_serve_dashboard_command
from solar_forecast.cli.job_commands import handle_status_command
from solar_forecast.cli.training_commands import handle_train_command
from solar_forecast.cli.verification_commands import handle_verify_e2e_command


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Solar forecast and anomaly monitoring")
    commands = parser.add_subparsers(dest="command", required=True)

    pipeline = commands.add_parser("pipeline", help="Run preprocessing, CNN training, analysis, and reporting")
    pipeline.add_argument("target")
    pipeline.add_argument("--data")
    pipeline.add_argument("--input-dir", default=str(LEGACY_MERGED_ROOT))
    pipeline.add_argument("--features")
    pipeline.add_argument("--output-dir", default=str(PIPELINE_RUNS_ROOT))
    pipeline.add_argument("--epochs", type=int, default=50)
    pipeline.add_argument("--n-trials", type=int, default=10)
    pipeline.add_argument("--optimizer-timeout-seconds", type=int)
    pipeline.add_argument("--no-optuna", action="store_true")
    pipeline.add_argument("--reinforcement", action="store_true")
    pipeline.add_argument("--contamination", type=float, default=0.05)
    pipeline.add_argument("--artifact-level", choices=["minimal", "standard", "debug"], default="minimal")
    add_sequence_arguments(pipeline)
    pipeline.set_defaults(func=handle_pipeline_command)

    hybrid = commands.add_parser("hybrid", help="Run the explainable dynamic hybrid")
    hybrid.add_argument("validation")
    hybrid.add_argument("test")
    hybrid.add_argument("--output-dir", default=str(HYBRID_EXPERIMENT_ROOT))
    hybrid.add_argument("--artifact-level", choices=["minimal", "standard", "debug"], default="minimal")
    hybrid.set_defaults(func=handle_hybrid_command)

    collect = commands.add_parser("collect", help="Collect official generation and KMA data")
    collect.add_argument("--start-date", required=True)
    collect.add_argument("--end-date")
    collect.add_argument("--sources", default="koen,kospo,ewp,iwest,kma")
    collect.add_argument("--output-dir", default=str(BRONZE_ROOT))
    collect.add_argument("--standardized-output-dir", default=str(COLLECTOR_SILVER_ROOT))
    collect.add_argument(
        "--download-date",
        help="Override the YYYY-MM-DD collection date embedded in canonical filenames",
    )
    collect.add_argument("--overwrite", action="store_true")
    collect.add_argument("--komipo-station-codes")
    collect.add_argument("--api-max-calls", type=int, default=900)
    collect.add_argument("--station-ids", help="Comma-separated ASOS station IDs; required in API mode")
    collect.add_argument("--weather-root", default=str(WEATHER_ARCHIVE_ROOT))
    collect.add_argument("--kma-mode", choices=["auto", "api", "browser"], default="auto",
                         help="auto uses KMA_ASOS_SERVICE_KEY when set; otherwise the existing browser collector")
    collect.set_defaults(func=handle_collect_command)

    prepare = commands.add_parser(
        "prepare-data",
        help="Standardize retained public-provider files and build leakage-safe model features",
    )
    prepare.add_argument("--input-root", default=str(GENERATION_ARCHIVE_ROOT))
    prepare.add_argument("--weather-root", default=str(WEATHER_ARCHIVE_ROOT))
    prepare.add_argument("--merged-source", default=str(LEGACY_MERGED_SOURCE))
    prepare.add_argument("--output-dir", default=str(SILVER_ROOT))
    prepare.add_argument(
        "--collected-generation-dir",
        default=str(COLLECTOR_SILVER_ROOT),
        help="Collector Silver files to admit into the Gold registry when they match the plant-hour schema",
    )
    prepare.add_argument(
        "--no-collected-downloads",
        action="store_true",
        help="Ignore collector Silver files and rebuild Gold only from retained historical archives",
    )
    prepare.set_defaults(func=handle_prepare_data_command)

    audit_candidate = commands.add_parser(
        "audit-candidate-data",
        help="Normalize and quality-gate staged additional generation data",
    )
    audit_candidate.add_argument(
        "--source-dir",
        default="file/raw/한국농어촌공사/영암",
    )
    audit_candidate.add_argument(
        "--output-dir",
        default="file/standardized/candidates/krc_yeongam",
    )
    audit_candidate.set_defaults(func=handle_audit_candidate_data_command)

    train = commands.add_parser("train", help="Run one independently locked model job")
    train.add_argument("model", choices=["xgboost", "cnn_bilstm"])
    train.add_argument("--config")
    train.add_argument("--smoke", action="store_true")
    train.add_argument(
        "--no-optuna",
        action="store_true",
        help="Skip hyperparameter search and use the fixed model config",
    )
    train.add_argument(
        "--max-trials",
        type=int,
        help="Override the maximum total trials in the resumable study",
    )
    train.add_argument(
        "--optimizer-timeout-seconds",
        type=int,
        help="Override the wall-time budget for this optimization call",
    )
    train.set_defaults(func=handle_train_command)

    verify_e2e = commands.add_parser(
        "verify-e2e",
        help="Run a bounded collect/prepare/train/dashboard wiring verification",
    )
    verify_e2e.add_argument(
        "--collect",
        action="store_true",
        help="Include official-source collection before local preprocessing",
    )
    verify_e2e.add_argument(
        "--start-date",
        help="Collection start date; required only when --collect is used",
    )
    verify_e2e.add_argument("--end-date")
    verify_e2e.add_argument("--sources", default="koen,kospo,ewp,iwest,kma")
    verify_e2e.add_argument("--collection-output-dir", default=str(BRONZE_ROOT))
    verify_e2e.add_argument(
        "--collection-standardized-output-dir",
        default=str(COLLECTOR_SILVER_ROOT),
    )
    verify_e2e.add_argument("--download-date")
    verify_e2e.add_argument("--overwrite", action="store_true")
    verify_e2e.add_argument("--komipo-station-codes")
    verify_e2e.add_argument("--api-max-calls", type=int, default=900)
    verify_e2e.add_argument("--station-ids", help="Comma-separated ASOS station IDs for --collect")
    verify_e2e.add_argument("--kma-mode", choices=["auto", "api", "browser"], default="auto")
    verify_e2e.add_argument(
        "--allow-collection-failures",
        action="store_true",
        help="Continue local verification if an external source is temporarily unavailable",
    )
    verify_e2e.add_argument(
        "--collection-manifest",
        help="Existing collection manifest to inspect when not rerunning downloads",
    )
    verify_e2e.add_argument("--skip-prepare", action="store_true")
    verify_e2e.add_argument("--input-root", default=str(GENERATION_ARCHIVE_ROOT))
    verify_e2e.add_argument("--weather-root", default=str(WEATHER_ARCHIVE_ROOT))
    verify_e2e.add_argument("--merged-source", default=str(LEGACY_MERGED_SOURCE))
    verify_e2e.add_argument("--output-dir", default=str(SILVER_ROOT))
    verify_e2e.add_argument(
        "--collected-generation-dir",
        default=str(COLLECTOR_SILVER_ROOT),
        help="Collector Silver files to admit into Gold during prepare-data",
    )
    verify_e2e.add_argument(
        "--no-collected-downloads",
        action="store_true",
        help="Run prepare-data without collector Silver admission",
    )
    verify_e2e.add_argument("--models", default="xgboost,cnn_bilstm")
    verify_e2e.add_argument(
        "--full-train",
        action="store_true",
        help="Run full model training instead of default smoke wiring checks",
    )
    verify_e2e.add_argument("--no-optuna", action="store_true")
    verify_e2e.add_argument("--max-trials", type=int)
    verify_e2e.add_argument("--optimizer-timeout-seconds", type=int)
    verify_e2e.add_argument("--skip-dashboard", action="store_true")
    verify_e2e.add_argument("--dashboard-output-dir", default="dashboard")
    verify_e2e.add_argument("--report-root", default=str(VERIFICATION_ROOT))
    verify_e2e.set_defaults(func=handle_verify_e2e_command)

    status = commands.add_parser("status", help="Show the active training job")
    status.set_defaults(func=handle_status_command)

    jobs = commands.add_parser(
        "jobs",
        help="List runnable job boundaries and future worker split decisions",
    )
    jobs.add_argument(
        "--json",
        action="store_true",
        help="Print the full job contract catalog as JSON",
    )
    jobs.set_defaults(func=handle_jobs_command)

    job_contract = commands.add_parser(
        "job-contract",
        help="Print one job's input/output artifact contract as JSON",
    )
    job_contract.add_argument(
        "job",
        choices=[contract.job_id for contract in list_job_contracts()],
    )
    job_contract.set_defaults(func=handle_job_contract_command)

    evaluate_features = commands.add_parser(
        "evaluate-features",
        help="Compare feature contracts with purged rolling-origin validation",
    )
    evaluate_features.add_argument("--data", default=str(GOLD_DATASET_PATH))
    evaluate_features.add_argument("--output-dir", default=str(FEATURE_EVALUATION_ROOT))
    evaluate_features.add_argument("--folds", type=int, default=3)
    evaluate_features.add_argument("--validation-window-hours", type=int, default=2160)
    evaluate_features.add_argument("--calibration-fraction", type=float, default=0.10)
    evaluate_features.add_argument("--test-fraction", type=float, default=0.15)
    evaluate_features.add_argument("--gap-hours", type=int, default=168)
    evaluate_features.add_argument("--n-estimators", type=int, default=300)
    evaluate_features.set_defaults(func=handle_evaluate_features_command)

    notify = commands.add_parser(
        "notify-anomalies",
        help="Validate operational anomaly JSONL and enqueue Kakao/SMS notifications",
    )
    notify.add_argument("--events", required=True)
    notify.add_argument(
        "--manifest",
        help="Integrity manifest; defaults to events.manifest.json beside --events",
    )
    notify.add_argument(
        "--routes",
        required=True,
        help="Runtime route directory containing phone environment-variable names only",
    )
    notify.add_argument("--outbox")
    notify.add_argument(
        "--template-contract",
        default="solar-anomaly-v1",
        help="Internal version of the approved notification message contract",
    )
    notify.add_argument(
        "--live",
        action="store_true",
        help="Allow external delivery only when the two environment safety gates also pass",
    )
    notify.set_defaults(func=handle_notify_anomalies_command)

    dashboard = commands.add_parser(
        "build-dashboard",
        help="Refresh dashboard data from current registry, quality, and model manifests",
    )
    dashboard.add_argument("--output-dir", default="dashboard")
    dashboard.set_defaults(func=handle_build_dashboard_command)

    prepare_boundaries = commands.add_parser(
        "prepare-boundaries",
        help="Convert an official SGIS province Shapefile to validated dashboard GeoJSON",
    )
    prepare_boundaries.add_argument("--source-shp", required=True)
    prepare_boundaries.add_argument("--source-archive", required=True)
    prepare_boundaries.add_argument("--output", default="map/json/geoJson.json")
    prepare_boundaries.add_argument("--reference-date", required=True)
    prepare_boundaries.add_argument("--archive-sha256", required=True)
    prepare_boundaries.add_argument("--shapefile-sha256", required=True)
    prepare_boundaries.add_argument(
        "--provider", default="국가데이터처 통계지리정보서비스(SGIS)"
    )
    prepare_boundaries.add_argument(
        "--source-url",
        default="https://www.data.go.kr/data/15129688/fileData.do",
    )
    prepare_boundaries.add_argument("--simplify-meters", type=float, default=150.0)
    prepare_boundaries.add_argument("--precision", type=int, default=6)
    prepare_boundaries.set_defaults(func=handle_prepare_boundaries_command)

    serve_dashboard = commands.add_parser(
        "serve-dashboard",
        help="Refresh and serve the dashboard with the correct document root",
    )
    serve_dashboard.add_argument("--host", default="127.0.0.1")
    serve_dashboard.add_argument("--port", type=int, default=5500)
    serve_dashboard.add_argument("--output-dir", default="dashboard")
    serve_dashboard.add_argument(
        "--no-refresh",
        action="store_true",
        help="Serve the existing dashboard payload without rebuilding it",
    )
    serve_dashboard.set_defaults(func=handle_serve_dashboard_command)
    return parser
