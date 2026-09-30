from __future__ import annotations

from datetime import date
import json

from solar_forecast.infrastructure.artifact_store import write_manifest
from solar_forecast.cli import build_parser
from solar_forecast.cli import main
from solar_forecast.collectors import CollectionConfig
from solar_forecast.collectors import CollectionService
from solar_forecast.jobs.contracts import COLLECTION_MANIFEST_CONTRACT
from solar_forecast.jobs.contracts import JOB_CONTRACT_SCHEMA_VERSION
from solar_forecast.jobs.contracts import MODEL_READY_MANIFEST_CONTRACT
from solar_forecast.jobs.contracts import TRAINING_RUN_MANIFEST_CONTRACT
from solar_forecast.jobs.contracts import job_contract_catalog
from solar_forecast.jobs.contracts import list_job_contracts


def test_job_contract_catalog_documents_modular_monolith_decision() -> None:
    catalog = job_contract_catalog()

    assert catalog["schema_version"] == JOB_CONTRACT_SCHEMA_VERSION
    assert "modular monolith" in catalog["architecture_decision"]
    assert {job.job_id for job in list_job_contracts()} >= {
        "collect",
        "prepare-data",
        "train",
        "build-dashboard",
        "notify-anomalies",
        "verify-e2e",
    }


def test_job_contract_cli_prints_worker_ready_boundary(capsys) -> None:
    main(["job-contract", "notify-anomalies"])

    payload = json.loads(capsys.readouterr().out)
    assert payload["job_id"] == "notify-anomalies"
    assert payload["worker_ready"] is True
    assert "external" in payload["worker_profile"]
    assert payload["inputs"][0]["contract"] == "solar-anomaly-event-manifest.v1"


def test_jobs_cli_lists_boundaries() -> None:
    parser = build_parser()

    args = parser.parse_args(["jobs", "--json"])
    assert args.json is True
    assert args.func.__name__ == "handle_jobs_command"


def test_training_manifest_includes_versioned_contract(tmp_path) -> None:
    manifest = write_manifest(
        tmp_path / "manifest.json",
        status="running",
        model="xgboost",
        run_id="run-1",
        details={"run_context": {"execution_mode": "smoke"}},
    )

    payload = json.loads(manifest.read_text(encoding="utf-8"))
    assert payload["contract"] == TRAINING_RUN_MANIFEST_CONTRACT
    assert payload["schema_version"] == JOB_CONTRACT_SCHEMA_VERSION


def test_collection_manifest_includes_versioned_contract(tmp_path) -> None:
    CollectionService(
        CollectionConfig(
            start_date=date(2025, 1, 1),
            end_date=date(2025, 1, 1),
            sources=("unknown",),
            output_dir=tmp_path / "raw",
            standardized_output_dir=tmp_path / "standardized" / "downloads",
        )
    ).run()

    manifest_path = next((tmp_path / "raw" / "runs").glob("*/collection_manifest.json"))
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert payload["contract"] == COLLECTION_MANIFEST_CONTRACT
    assert payload["schema_version"] == JOB_CONTRACT_SCHEMA_VERSION


def test_prepare_data_contract_declares_gold_model_ready_schema() -> None:
    prepare = next(job for job in list_job_contracts() if job.job_id == "prepare-data")

    assert prepare.outputs[0].contract == MODEL_READY_MANIFEST_CONTRACT
    assert prepare.outputs[0].schema["properties"]["target"]["const"] == "generation_mwh"
    assert "schema_version" in prepare.outputs[0].schema["required"]
