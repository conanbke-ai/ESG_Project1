"""Job commands: CLI input translation and command dispatch."""
from __future__ import annotations

import argparse
import json
from solar_forecast.jobs.contracts import get_job_contract, job_contract_catalog, list_job_contracts
from solar_forecast.config_loader import PROJECT_ROOT


def handle_status_command(_: argparse.Namespace) -> None:
    lock = PROJECT_ROOT / "artifacts" / ".training.lock"
    print(f"Training active: {lock.read_text(encoding='utf-8')}" if lock.exists() else "No training job is active")


def handle_jobs_command(args: argparse.Namespace) -> None:
    if args.json:
        print(json.dumps(job_contract_catalog(), ensure_ascii=False, indent=2))
        return
    print("Runnable job boundaries")
    for contract in list_job_contracts():
        readiness = "worker-ready" if contract.worker_ready else "readiness-check"
        outputs = ", ".join(output.name for output in contract.outputs) or "none"
        print(f"- {contract.job_id} [{readiness}]")
        print(f"  command: {contract.command}")
        print(f"  outputs: {outputs}")
        print(f"  split: {contract.split_decision}")


def handle_job_contract_command(args: argparse.Namespace) -> None:
    try:
        contract = get_job_contract(args.job)
    except KeyError as exc:
        raise SystemExit(str(exc)) from exc
    print(json.dumps(contract.as_dict(), ensure_ascii=False, indent=2))
