"""Read-only config/schema and checkpoint metadata compatibility inspection.

No optimizer/study is opened and no torch/pickle checkpoint is deserialized.
A matching training fingerprint is necessary, never sufficient for resuming a
specific stage: the trainer must additionally validate its stage signature.
"""
from __future__ import annotations

import json
from pathlib import Path

from solar_forecast.config_loader import ModelJobConfig, PROJECT_ROOT
from solar_forecast.datasets.numeric_preprocessor import require_model_quality_filter
from solar_forecast.infrastructure.artifact_store import sha256_file
from solar_forecast.models.shared.checkpoint_store import (
    CHECKPOINT_CONTRACT, TrainingCheckpointStore,
)


def audit_training_compatibility(config: ModelJobConfig, gold: dict, *, project_root: Path = PROJECT_ROOT) -> dict:
    """Use the unchanged training fingerprint algorithm and saved sidecars."""
    root = Path(project_root).resolve()
    values = config.values
    source = Path(values["input_dataset"])
    source = (source if source.is_absolute() else root / source).resolve()
    audited = {Path(gold["dataset"]["path"]).resolve()}
    declared_parts = {Path(p["path"]).resolve() for p in gold["partitions"]}
    # A config must target the audited dataset or exactly its audited partition
    # inventory; never classify arbitrary paths with a matching header as Gold.
    if source.is_dir():
        from solar_forecast.datasets.repository import DatasetRepository
        same_source = set(DatasetRepository.training_files(source)) == declared_parts
    else:
        same_source = source in audited
    columns = gold["dataset"]["columns"]
    features = list(values.get("feature_columns") or [])
    quality = require_model_quality_filter(values)
    required = ["plant_id", "timestamp", "region", "plant", values["target_column"], quality, *features]
    missing = sorted(set(required).difference(columns))
    issues = []
    if not same_source:
        issues.append("config_does_not_reference_audited_gold")
    if missing:
        issues.append("required_columns_missing")
    if not features or len(features) != len(set(features)):
        issues.append("feature_list_empty_or_duplicate")
    if values["target_column"] != gold["target"]:
        issues.append("target_contract_mismatch")
    return {
        **inspect_checkpoint_metadata(config, project_root=root),
        "schema_status": "FAIL" if issues else "PASS",
        "schema_issues": issues,
        "missing_columns": missing,
        "ordered_features": features,
        "source": str(source),
        "quality_filter": quality,
        "energy_source_filter": values.get("energy_source_filter"),
        "eligible_rows_before_horizon_alignment": gold["dataset"]["eligible_rows"].get(values.get("energy_source_filter")),
    }


def inspect_checkpoint_metadata(config: ModelJobConfig, *, project_root: Path = PROJECT_ROOT) -> dict:
    """Fingerprint the configured sources and inspect saved metadata and payload presence."""
    if config.model not in {"xgboost", "cnn_bilstm"}:
        raise ValueError("Checkpoint metadata inspection supports XGBoost and CNN-BiLSTM configs")
    # Path spelling is part of model_values in the existing identity.
    root = Path(project_root).resolve()
    store = TrainingCheckpointStore.from_config(config, project_root=root)
    fingerprint = store.fingerprint
    records = []
    model_root = store.root / config.model
    for path in sorted(model_root.glob("*/*.meta.json")):
        meta = json.loads(path.read_text(encoding="utf-8-sig"))
        payload = path.with_name(path.name.removesuffix(".meta.json"))
        problems = []
        for key, expected in (("checkpoint_contract", CHECKPOINT_CONTRACT), ("model", config.model),
                              ("fingerprint", path.parent.name)):
            if meta.get(key) != expected:
                problems.append(key)
        if not payload.is_file():
            problems.append("payload_missing")
        if not meta.get("signature") or not meta.get("stage"):
            problems.append("stage_signature_missing")
        records.append({"metadata_path": str(path), "metadata_sha256": sha256_file(path),
                        "checkpoint_path": str(payload), "checkpoint_bytes": payload.stat().st_size if payload.is_file() else None,
                        "checkpoint_contract": meta.get("checkpoint_contract"),
                        "fingerprint": meta.get("fingerprint"), "completed": meta.get("completed"),
                        "progress": meta.get("progress"), "stage": meta.get("stage"),
                        "same_training_fingerprint": meta.get("fingerprint") == fingerprint,
                        "metadata_issues": problems})
    matches = [r for r in records if r["same_training_fingerprint"] and not r["metadata_issues"]]
    return {"model": config.model, "config": str(config.source), "config_sha256": sha256_file(config.source),
            "training_fingerprint": fingerprint, "checkpoint_directory": str(store.directory),
            "checkpoint_scope": "configured_fingerprint_and_saved_metadata_only",
            "checkpoint_status": "MATCHING_METADATA_REQUIRES_STAGE_SIGNATURE" if matches else "NO_EXACT_TRAINING_IDENTITY",
            "checkpoints": records, "matching_metadata_count": len(matches),
            "unsupported_checkpoint_contract_count": sum(r["checkpoint_contract"] != CHECKPOINT_CONTRACT for r in records),
            "checkpoint_payload_loaded": False, "resume_executed": False,
            "resume_policy": "unchanged trainer validates contract/model/fingerprint/stage/signature before resume"}
