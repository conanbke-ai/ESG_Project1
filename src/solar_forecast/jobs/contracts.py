"""독립 실행 job 목록과 교환 manifest의 버전·입출력 계약 정의."""
from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any


JOB_CONTRACT_SCHEMA_VERSION = 1

COLLECTION_MANIFEST_CONTRACT = "solar-collection-manifest.v1"
MODEL_READY_MANIFEST_CONTRACT = "solar-model-ready-manifest.v1"
TRAINING_RUN_MANIFEST_CONTRACT = "solar-training-run-manifest.v1"
E2E_VERIFICATION_MANIFEST_CONTRACT = "solar-e2e-verification.v1"
ANOMALY_EVENT_MANIFEST_CONTRACT = "solar-anomaly-event-manifest.v1"


@dataclass(frozen=True)
class ArtifactContract:
    """A versioned artifact exchanged between independently runnable jobs."""

    name: str
    contract: str
    path_pattern: str
    schema: dict[str, Any]
    required_for_next_job: bool = True
    notes: tuple[str, ...] = ()

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class JobContract:
    """Public job boundary that can later become a worker/container entrypoint."""

    job_id: str
    command: str
    purpose: str
    worker_profile: str
    side_effects: tuple[str, ...]
    inputs: tuple[ArtifactContract, ...]
    outputs: tuple[ArtifactContract, ...]
    worker_ready: bool
    split_decision: str
    notes: tuple[str, ...] = ()

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


def _manifest_schema(
    contract: str,
    *,
    required: tuple[str, ...],
    properties: dict[str, Any],
    contract_field: str = "contract",
) -> dict[str, Any]:
    schema_properties: dict[str, Any] = {
        contract_field: {"const": contract},
        **properties,
    }
    if contract_field == "contract":
        schema_properties["schema_version"] = {"const": JOB_CONTRACT_SCHEMA_VERSION}
    required_fields = [contract_field, *required]
    if contract_field == "contract" and "schema_version" not in required_fields:
        required_fields.insert(1, "schema_version")
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "required": required_fields,
        "properties": schema_properties,
        "additionalProperties": True,
    }


COLLECTION_MANIFEST_SCHEMA = _manifest_schema(
    COLLECTION_MANIFEST_CONTRACT,
    required=("started_at", "start_date", "end_date", "results", "file_artifacts"),
    properties={
        "started_at": {"type": "string"},
        "start_date": {"type": "string", "format": "date"},
        "end_date": {"type": "string", "format": "date"},
        "results": {
            "type": "array",
            "items": {
                "type": "object",
                "required": ("source", "status", "rows", "files", "message"),
                "properties": {
                    "source": {"type": "string"},
                    "status": {"type": "string"},
                    "rows": {"type": "integer"},
                    "files": {"type": "array", "items": {"type": "string"}},
                    "message": {"type": "string"},
                },
                "additionalProperties": True,
            },
        },
        "file_artifacts": {"type": "array"},
    },
)

MODEL_READY_MANIFEST_SCHEMA = _manifest_schema(
    MODEL_READY_MANIFEST_CONTRACT,
    required=(
        "created_at",
        "dataset",
        "partitioned_dataset",
        "target",
        "features",
        "training_eligibility",
        "rows",
        "plants",
        "start",
        "end",
    ),
    properties={
        "created_at": {"type": "string"},
        "dataset": {"type": "string"},
        "partitioned_dataset": {"type": "object"},
        "target": {"const": "generation_mwh"},
        "features": {"type": "array", "items": {"type": "string"}},
        "training_eligibility": {"type": "object"},
        "rows": {"type": "integer"},
        "plants": {"type": "integer"},
        "start": {"type": "string"},
        "end": {"type": "string"},
    },
)

TRAINING_RUN_MANIFEST_SCHEMA = _manifest_schema(
    TRAINING_RUN_MANIFEST_CONTRACT,
    required=("status", "model", "run_id", "updated_at_utc", "details"),
    properties={
        "status": {"enum": ["running", "completed", "failed"]},
        "model": {"enum": ["xgboost", "cnn_bilstm"]},
        "run_id": {"type": "string"},
        "updated_at_utc": {"type": "string"},
        "details": {"type": "object"},
    },
)

E2E_VERIFICATION_MANIFEST_SCHEMA = _manifest_schema(
    E2E_VERIFICATION_MANIFEST_CONTRACT,
    required=("status", "started_at_utc", "finished_at_utc", "execution", "summary", "steps"),
    properties={
        "status": {"enum": ["passed", "passed_with_warnings", "failed"]},
        "started_at_utc": {"type": "string"},
        "finished_at_utc": {"type": "string"},
        "execution": {"type": "object"},
        "summary": {"type": "object"},
        "steps": {"type": "array"},
    },
)

ANOMALY_EVENT_MANIFEST_SCHEMA = _manifest_schema(
    ANOMALY_EVENT_MANIFEST_CONTRACT,
    contract_field="manifest_contract",
    required=("event_contract", "scope", "events_file", "sha256", "event_count", "run_id", "detector_version"),
    properties={
        "event_contract": {"const": "solar-anomaly-event.v1"},
        "scope": {"const": "operational"},
        "events_file": {"type": "string"},
        "sha256": {"type": "string"},
        "event_count": {"type": "integer"},
        "run_id": {"type": "string"},
        "detector_version": {"type": "string"},
    },
)


def _artifact(
    name: str,
    contract: str,
    path_pattern: str,
    schema: dict[str, Any],
    *,
    required_for_next_job: bool = True,
    notes: tuple[str, ...] = (),
) -> ArtifactContract:
    return ArtifactContract(
        name=name,
        contract=contract,
        path_pattern=path_pattern,
        schema=schema,
        required_for_next_job=required_for_next_job,
        notes=notes,
    )


JOB_CONTRACTS: tuple[JobContract, ...] = (
    JobContract(
        job_id="benchmark",
        command="python app.py benchmark --config config/experiments/optimized.json [--plan]",
        purpose="실측 기반 예측 기간별로 모델을 독립 최적화하고, Test 이전에 하이브리드 채택을 결정합니다.",
        worker_profile="sequential training and evaluation worker; artifact contracts; no network calls",
        side_effects=("artifacts/benchmarks/<run-id>/에 설정, 후보 모델, 선택 근거와 예측을 기록합니다.",),
        inputs=(_artifact("experiment", "solar-optimized-experiment.v1", "config/experiments/optimized.json",
                          {"type": "object", "required": ["contract", "models", "horizons_hours", "input_dataset"]}),),
        outputs=(_artifact("benchmark", "solar-optimized-benchmark.v1", "artifacts/benchmarks/<run-id>/manifest.json",
                           {"type": "object", "required": ["contract", "status", "execution_mode", "tasks", "provenance"]}),),
        worker_ready=True,
        split_decision="같은 코드베이스의 독립 실행 job으로 유지하며, 완료된 산출물만 대시보드에 전달합니다.",
        notes=("예보 API를 요구하지 않습니다. Test 점수는 모델 선택에 사용하지 않습니다.",),
    ),
    JobContract(
        job_id="collect",
        command="python app.py collect --start-date <YYYY-MM-DD> [--end-date <YYYY-MM-DD>]",
        purpose="공식 발전량/KMA 원본을 Bronze로 보존하고, 표준화 가능한 발전량 파일은 Silver CSV로 생성합니다.",
        worker_profile="io-bound collector; external read; resumable by source/date output paths",
        side_effects=(
            "공식 웹/API에서 파일을 다운로드합니다.",
            "file/raw/<source>/와 file/standardized/downloads/<source>/에 산출물을 씁니다.",
            "ASOS API/browser 관측값은 file/KMA_data_file/OBS_ASOS_TIM_<year>.csv로 병합합니다.",
        ),
        inputs=(),
        outputs=(
            _artifact(
                "collection_manifest",
                COLLECTION_MANIFEST_CONTRACT,
                "file/raw/runs/<run-id>/collection_manifest.json",
                COLLECTION_MANIFEST_SCHEMA,
            ),
        ),
        worker_ready=True,
        split_decision="독립 worker/container로 분리 가능하지만 현재는 같은 CLI의 collect job으로 유지합니다.",
    ),
    JobContract(
        job_id="prepare-data",
        command="python app.py prepare-data",
        purpose="보관 원본과 collector Silver를 registry/admission으로 심사한 뒤 Gold model-ready 파티션을 만듭니다.",
        worker_profile="cpu/io-bound dataset builder; no external delivery; deterministic artifacts",
        side_effects=(
            "file/standardized/generation/ 파티션을 갱신합니다.",
            "file/standardized/model_ready.csv.gz와 model_ready_parts/를 갱신합니다.",
        ),
        inputs=(
            _artifact(
                "collection_manifest",
                COLLECTION_MANIFEST_CONTRACT,
                "file/raw/runs/<run-id>/collection_manifest.json",
                COLLECTION_MANIFEST_SCHEMA,
                required_for_next_job=False,
                notes=("prepare-data는 manifest가 없어도 기존 file/standardized/downloads를 재심사할 수 있습니다.",),
            ),
        ),
        outputs=(
            _artifact(
                "model_ready_manifest",
                MODEL_READY_MANIFEST_CONTRACT,
                "file/standardized/model_ready_manifest.json",
                MODEL_READY_MANIFEST_SCHEMA,
            ),
        ),
        worker_ready=True,
        split_decision="마이크로서비스가 아니라 Gold dataset builder job으로 유지합니다.",
    ),
    JobContract(
        job_id="train",
        command="python app.py train <xgboost|cnn_bilstm>",
        purpose="동일한 Gold 계약에서 모델별 전략을 실행하고 checkpoint와 평가 예측을 남깁니다.",
        worker_profile="cpu/gpu-bound training worker; exclusive lock; resumable checkpoint",
        side_effects=(
            "artifacts/models/<model>/<run-id>/에 모델·예측·manifest를 씁니다.",
            "artifacts/checkpoints/<model>/<fingerprint>/에 재개 상태를 씁니다.",
        ),
        inputs=(
            _artifact(
                "model_ready_manifest",
                MODEL_READY_MANIFEST_CONTRACT,
                "file/standardized/model_ready_manifest.json",
                MODEL_READY_MANIFEST_SCHEMA,
            ),
        ),
        outputs=(
            _artifact(
                "training_manifest",
                TRAINING_RUN_MANIFEST_CONTRACT,
                "artifacts/models/<model>/<run-id>/manifest.json",
                TRAINING_RUN_MANIFEST_SCHEMA,
            ),
        ),
        worker_ready=True,
        split_decision="모델별 마이크로서비스 대신 같은 TrainingService의 strategy job으로 유지합니다.",
    ),
    JobContract(
        job_id="build-dashboard",
        command="python app.py build-dashboard",
        purpose="전국 설비 현황과 정식 모델 분석 산출물을 정적 dashboard payload로 projection합니다.",
        worker_profile="report builder; read-heavy; static artifact publisher",
        side_effects=("dashboard/data/dashboard_data.json을 원자적으로 갱신합니다.",),
        inputs=(
            _artifact(
                "model_ready_manifest",
                MODEL_READY_MANIFEST_CONTRACT,
                "file/standardized/model_ready_manifest.json",
                MODEL_READY_MANIFEST_SCHEMA,
                required_for_next_job=False,
            ),
            _artifact(
                "training_manifest",
                TRAINING_RUN_MANIFEST_CONTRACT,
                "artifacts/models/<model>/<run-id>/manifest.json",
                TRAINING_RUN_MANIFEST_SCHEMA,
                required_for_next_job=False,
            ),
        ),
        outputs=(),
        worker_ready=True,
        split_decision="사용자용 정적 projection job으로 유지하며 별도 웹 마이크로서비스는 아직 만들지 않습니다.",
    ),
    JobContract(
        job_id="notify-anomalies",
        command="python app.py notify-anomalies --events <events.jsonl> --routes <routes.json> [--live]",
        purpose="운영 이상 이벤트 manifest를 검증하고 카카오/SMS outbox를 dispatch합니다.",
        worker_profile="external side-effect dispatcher; transactional outbox; fail-closed live gate",
        side_effects=(
            "dry-run에서는 artifacts/notifications/dry_run.sqlite3만 사용합니다.",
            "live 승인 시 외부 Kakao/SMS provider로 메시지를 접수합니다.",
        ),
        inputs=(
            _artifact(
                "operational_anomaly_event_manifest",
                ANOMALY_EVENT_MANIFEST_CONTRACT,
                "artifacts/anomalies/<run-id>/events.manifest.json",
                ANOMALY_EVENT_MANIFEST_SCHEMA,
            ),
        ),
        outputs=(),
        worker_ready=True,
        split_decision="외부 부작용이 있으므로 실제 분리 1순위입니다. 현재는 dispatcher job 객체로 CLI에서 격리합니다.",
    ),
    JobContract(
        job_id="verify-e2e",
        command="python app.py verify-e2e [--collect --start-date <YYYY-MM-DD>]",
        purpose="collect/prepare/train/dashboard의 배선과 산출물 계약을 bounded smoke로 검증합니다.",
        worker_profile="orchestration check; smoke by default; no formal metric pollution",
        side_effects=("artifacts/verification/e2e/<run-id>/verification_report.json을 씁니다.",),
        inputs=(),
        outputs=(
            _artifact(
                "verification_report",
                E2E_VERIFICATION_MANIFEST_CONTRACT,
                "artifacts/verification/e2e/<run-id>/verification_report.json",
                E2E_VERIFICATION_MANIFEST_SCHEMA,
            ),
        ),
        worker_ready=False,
        split_decision="운영 서비스가 아니라 release/readiness check로 유지합니다.",
    ),
)


def list_job_contracts() -> tuple[JobContract, ...]:
    return JOB_CONTRACTS


def get_job_contract(job_id: str) -> JobContract:
    for contract in JOB_CONTRACTS:
        if contract.job_id == job_id:
            return contract
    known = ", ".join(contract.job_id for contract in JOB_CONTRACTS)
    raise KeyError(f"Unknown job contract '{job_id}'. Known jobs: {known}")


def job_contract_catalog() -> dict[str, Any]:
    return {
        "schema_version": JOB_CONTRACT_SCHEMA_VERSION,
        "architecture_decision": (
            "Keep a modular monolith now, expose independently runnable CLI jobs, "
            "and split only the side-effect notification dispatcher first when deployment needs justify it."
        ),
        "jobs": [contract.as_dict() for contract in JOB_CONTRACTS],
    }


__all__ = [
    "ANOMALY_EVENT_MANIFEST_CONTRACT",
    "COLLECTION_MANIFEST_CONTRACT",
    "E2E_VERIFICATION_MANIFEST_CONTRACT",
    "JOB_CONTRACT_SCHEMA_VERSION",
    "MODEL_READY_MANIFEST_CONTRACT",
    "TRAINING_RUN_MANIFEST_CONTRACT",
    "ArtifactContract",
    "JobContract",
    "get_job_contract",
    "job_contract_catalog",
    "list_job_contracts",
]
