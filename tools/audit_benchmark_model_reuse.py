"""Compare a saved benchmark with current training contracts without fitting models."""

from __future__ import annotations

import argparse
from collections import Counter
from contextlib import closing
import csv
import json
from pathlib import Path, PureWindowsPath
import re
from shutil import copyfile
import sqlite3
import subprocess
import sys
from tempfile import TemporaryDirectory

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))

from solar_forecast.config_loader import ModelJobConfig
from solar_forecast.evaluation.experiment_config import build_candidate_configs, load_experiment_config
from solar_forecast.features.future_weather import future_weather_feature_columns
from solar_forecast.infrastructure.artifact_store import sha256_file
from solar_forecast.models.shared.checkpoint_store import (
    CHECKPOINT_CONTRACT,
    CheckpointSettings,
    TrainingCheckpointStore,
    dataset_signature,
    stable_signature,
    training_identity_payload,
)
from solar_forecast.models.shared.optuna_study import OptimizationSettings

# The benchmark resolver uses this outer profile independently of the model
# profile stored inside resolved_config.json.
BENCHMARK_PROFILE = "historical_optimized"
SPLIT_FIELDS = ("train_end", "validation_end", "calibration_end", "test_end", "purge_gap_hours")
ARCHITECTURE_FIELDS = ("feature_columns", "sequence_length", "append_missing_indicators")
ABSENT = {"field_absent": True}


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def artifact_path(benchmark: Path, reference: str) -> Path:
    """Relocate the complete historical directory across Windows and Unix."""
    path = Path(reference)
    windows = PureWindowsPath(reference)
    if windows.is_absolute() or path.is_absolute():
        parts = windows.parts if windows.is_absolute() else path.parts
        if benchmark.name not in parts:
            raise ValueError(f"Evidence path is outside benchmark: {reference}")
        path = Path(*parts[parts.index(benchmark.name) + 1:])
    resolved = (benchmark / path).resolve()
    if not resolved.is_relative_to(benchmark) or not resolved.is_file():
        raise ValueError(f"Missing or external benchmark artifact: {reference}")
    return resolved


def stable_diff(before: dict, after: dict, prefix: str = "") -> dict:
    changes = {}
    for key in sorted(before.keys() | after.keys()):
        old, new = before.get(key, ABSENT), after.get(key, ABSENT)
        name = prefix + key
        if old == new:
            continue
        if isinstance(old, dict) and isinstance(new, dict) and ABSENT not in (old, new):
            changes.update(stable_diff(old, new, name + "."))
        else:
            changes[name] = {"historical": old, "current": new}
    return changes


def tracked_blob(path: Path) -> str | None:
    """Allow model loading only for bytes matching a committed Git blob."""
    commands = (
        ["rev-parse", "--show-toplevel"],
        ["hash-object", "--", str(path)],
    )
    root, actual = (subprocess.run(["git", "-C", str(path.parent), *cmd], capture_output=True, text=True)
                    for cmd in commands)
    if root.returncode or actual.returncode:
        return None
    relative = path.relative_to(root.stdout.strip()).as_posix()
    expected = subprocess.run(
        ["git", "-C", root.stdout.strip(), "rev-parse", "--verify", f"HEAD:{relative}"],
        capture_output=True,
        text=True,
    )
    oid = actual.stdout.strip()
    return oid if expected.returncode == 0 and oid == expected.stdout.strip() else None


def study_identity(config: ModelJobConfig, fingerprint: str) -> str:
    settings = OptimizationSettings.from_values(config.values, model=config.model).scoped(fingerprint)
    if config.model == "cnn_bilstm":
        from solar_forecast.models.cnn_bilstm.optimization import scope_cnn_readout_study

        settings = scope_cnn_readout_study(settings, config.values["optimizer"].get("search_space", {}))
    return settings.study_name


def historical_xgboost_stage(values: dict, details: dict, parameters: dict) -> str:
    """Reconstruct the 2026-09-22 stage from archived configuration and split."""
    params = {
        "tree_method": "hist",
        "eval_metric": "mae",
        "random_state": values["seed"],
        "n_jobs": values["n_jobs"],
        "n_estimators": values["optimizer"]["trial_max_estimators"],
        "early_stopping_rounds": values["optimizer"]["early_stopping_rounds"],
        **parameters,
    }
    return stable_signature({
        "params": params,
        "feature_columns": details["features"],
        "target_column": values["target_column"],
        "temporal_split": details["temporal_split"],
    })


def fitted_evidence(benchmark: Path, config: ModelJobConfig, details: dict, parameters: dict,
                    fingerprint: str, temporary_root: Path) -> dict:
    model = config.model
    model_path = artifact_path(benchmark, details["model_path" if model == "xgboost" else "checkpoint_path"])
    blob = tracked_blob(model_path)
    evidence = {
        "artifact": model_path.relative_to(benchmark).as_posix(),
        "sha256": sha256_file(model_path),
        "committed_git_blob": blob,
        "cpu_loaded": False,
    }
    if not blob:
        return {**evidence, "status": "unverified_model_bytes"}
    stage = details["checkpoint"]["stage"]
    split = details["temporal_split"]
    if model == "xgboost":
        from xgboost import XGBRegressor

        preprocessing = read_json(artifact_path(benchmark, details["preprocessing_path"]))
        regressor = XGBRegressor(device="cpu", n_jobs=1)
        regressor.load_model(model_path)
        booster = regressor.get_booster()
        if booster.feature_names != details["features"] or preprocessing["temporal_split"] != split:
            raise ValueError(f"Saved XGBoost preprocessing does not match manifest: {model_path}")
        signature = historical_xgboost_stage(config.values, details, parameters)
        if stage != f"final_fit_{signature[:20]}":
            raise ValueError("Reconstructed XGBoost stage differs from the manifest")
        # The fitted Booster embeds its original stage payload. Validate those
        # bytes with the actual checkpoint guard, in a disposable directory.
        store = TrainingCheckpointStore(temporary_root, model=model, fingerprint=fingerprint)
        copied = store.xgboost_path(stage)
        copied.parent.mkdir(parents=True, exist_ok=True)
        copyfile(model_path, copied)
        loaded = store.load_xgboost(stage, signature=signature)
        if loaded is None:
            raise ValueError("XGBoost artifact has no resumable checkpoint contract")
        _, state = loaded
        if not state["completed"] or state["progress"]["completed_rounds"] != details["checkpoint"]["completed_rounds"]:
            raise ValueError("Saved XGBoost progress does not match completed historical stage")
        evidence.update(
            architecture={"tree_method": "hist", "completed_rounds": state["progress"]["completed_rounds"]},
            scaler={"strategy": preprocessing["strategy"], "fitted_scaler": False},
            checkpoint_payload_validated_on_cpu=True,
        )
    else:
        import torch
        from solar_forecast.models.cnn_bilstm.input_preprocessing import validate_input_preprocessing
        from solar_forecast.models.cnn_bilstm.network import CnnBiLstmNetworkConfig, build_cnn_bilstm_network

        fitted = torch.load(model_path, weights_only=True, map_location="cpu")
        features = fitted["feature_columns"]
        preprocessing = validate_input_preprocessing(fitted["preprocessing"], features)
        network = build_cnn_bilstm_network(CnnBiLstmNetworkConfig(**fitted["config"]), device=torch.device("cpu"))
        network.load_state_dict(fitted["model_state"], strict=True)
        if features != details["features"] or preprocessing["temporal_split"] != split:
            raise ValueError(f"Saved CNN preprocessing does not match manifest: {model_path}")
        signature = stable_signature({
            "model_config": fitted["config"],
            "input_preprocessing_contract": config.values["input_preprocessing_contract"],
            "best_params": parameters,
            "epochs": config.values["epochs"],
            "train_sequences": split["counts"]["train"],
            "validation_sequences": split["counts"]["validation"],
        })
        evidence.update(
            architecture=fitted["config"],
            scaler={
                "contract": preprocessing["contract"],
                "strategy": preprocessing["strategy"],
                "fit_scope": preprocessing["fit_scope"],
                "ordered_features": features,
                "scaler_state_sha256": stable_signature({
                    key: preprocessing[key] for key in ("feature_medians", "feature_means", "feature_scales")
                }),
            },
            fitted_target_transform=fitted.get("target_transform", "identity"),
            checkpoint_payload_validated_on_cpu=False,
            fitted_artifact_has_optimizer_state="optimizer_state" in fitted,
        )
    if stage != f"final_fit_{signature[:20]}":
        raise ValueError(f"Reconstructed final-fit stage differs from manifest: {model_path}")
    return {
        **evidence,
        "historical_stage": stage,
        "reconstructed_historical_stage_signature": signature,
        "cpu_loaded": True,
    }


def read_only_studies(path: Path) -> dict[str, int]:
    if not path.is_file():
        return {}
    for suffix in ("-wal", "-journal"):
        journal = path.with_name(path.name + suffix)
        if journal.is_file() and journal.stat().st_size:
            raise ValueError(f"SQLite requires a quiescent snapshot: {path}")
    with closing(sqlite3.connect(path.resolve().as_uri() + "?mode=ro&immutable=1", uri=True)) as db:
        return dict(db.execute(
            "SELECT study_name, count(trial_id) FROM studies LEFT JOIN trials USING (study_id) GROUP BY study_name"
        ))


def historical_trial_evidence(benchmark: Path, details: dict, study: str, db_studies: dict) -> dict:
    optimizer = details["optimizer"]
    summary_path = artifact_path(benchmark, optimizer["summary_path"])
    trials_path = artifact_path(benchmark, optimizer["trials_path"])
    summary = read_json(summary_path)
    with trials_path.open(encoding="utf-8-sig", newline="") as stream:
        states = Counter(row["state"] for row in csv.DictReader(stream))
    finished = sum(states[state] for state in ("COMPLETE", "PRUNED", "FAIL"))
    if (summary["study_name"] != study or summary["selection_data"] != "validation_only"
            or summary["test_usage"] != "none" or summary["finished_trials"] != finished):
        raise ValueError(f"Study summary and reported identity disagree: {summary_path}")
    return {
        "study_name": study,
        "exported_trial_states": dict(states),
        "exported_finished_trials": finished,
        "historically_executed_trials_this_run": summary["executed_trials"],
        "sqlite_study_trial_rows": db_studies.get(study),
        "summary_sha256": sha256_file(summary_path),
        "trials_csv_sha256": sha256_file(trials_path),
    }


def current_comparisons(current: ModelJobConfig, historical: ModelJobConfig, dataset_digest: str,
                        current_dataset_digest: str | None) -> dict:
    old = training_identity_payload(historical, dataset_digest)
    new = training_identity_payload(current, dataset_digest)
    input_field = "input_dataset"
    old_path = old["model_values"][input_field]
    current_path = new["model_values"][input_field]
    path_only = {**old, "model_values": {**old["model_values"], input_field: current_path}}
    training_only = {**new, "model_values": {**new["model_values"], input_field: old_path}}
    archived_hash = stable_signature(old)
    current_configuration_hash = stable_signature(training_only)
    actual_hash = None
    if current_dataset_digest is not None:
        actual_hash = stable_signature(training_identity_payload(current, current_dataset_digest))
    values = current.values
    future = values["future_weather"]
    future_features = list(future_weather_feature_columns(future["feature_profile"])) if future["enabled"] else []
    from solar_forecast.models.cnn_bilstm.network import CnnBiLstmNetworkConfig

    return {
        "candidate_id": values["benchmark_candidate_id"],
        "effective_config_differences": stable_diff(old, new),
        "same_archived_dataset_digest_and_path_fingerprint": current_configuration_hash,
        "same_archived_dataset_digest_and_path_matches": current_configuration_hash == archived_hash,
        "path_only_fingerprint": stable_signature(path_only),
        "configured_paths_with_archived_digest_fingerprint": stable_signature(new),
        "actual_dataset_and_configuration_fingerprint": actual_hash,
        "current_study_name_with_archived_digest_and_path": study_identity(current, current_configuration_hash),
        "current_study_name_with_actual_dataset": study_identity(current, actual_hash) if actual_hash else None,
        "historical_features_and_lookback_equal": all(
            historical.values.get(key) == values.get(key) for key in ARCHITECTURE_FIELDS
        ),
        "current_future_features": future_features,
        "global_split_and_purge_equal": all(historical.values.get(key) == values.get(key) for key in SPLIT_FIELDS),
        "current_split": {key: values[key] for key in SPLIT_FIELDS},
        "current_scaler_policy": values["missing_value_policy"],
        "current_input_preprocessing_contract": values.get("input_preprocessing_contract"),
        "current_target_transform": values.get("target_transform", "identity"),
        "current_cnn_network_config_fields": list(CnnBiLstmNetworkConfig.__dataclass_fields__)
        if current.model == "cnn_bilstm" else None,
        "current_optimizer_search_space_equal":
            historical.values["optimizer"]["search_space"] == values["optimizer"]["search_space"],
        "current_final_fit_stage": "requires_effective_features_selected_parameters_and_current_split",
    }


def audit(benchmark: Path, experiment: Path, study_db: Path, checkpoint_root: Path) -> dict:
    benchmark = benchmark.resolve()
    manifest_path = benchmark / "manifest.json"
    manifest = read_json(manifest_path)
    if manifest["status"] != "completed" or manifest["execution_mode"] != "full":
        raise ValueError("Audit requires a completed historical full benchmark")
    provenance = manifest["provenance"]
    historical_digest = provenance["dataset_fingerprint"]
    current_experiment = load_experiment_config(experiment, project_root=ROOT)
    source = ROOT / current_experiment["input_dataset"]
    current_dataset_digest = dataset_signature(source) if source.exists() else None
    db_studies = read_only_studies(study_db)
    entries = []
    with TemporaryDirectory(prefix="solar-model-reuse-") as temporary:
        for task in manifest["tasks"]:
            horizon = task["horizon_hours"]
            if horizon not in current_experiment["horizons_hours"]:
                raise ValueError(f"Current experiment omits archived horizon: {horizon}")
            candidates = build_candidate_configs(current_experiment, horizon, Path("."), project_root=ROOT)
            for model, item in sorted(task["optimization"].items()):
                run_manifest = artifact_path(benchmark, item["run_dir"] + "/manifest.json")
                run = read_json(run_manifest)
                config_path = artifact_path(benchmark, item["resolved_config"])
                saved = ModelJobConfig(model, BENCHMARK_PROFILE, read_json(config_path), config_path)
                details = run["details"]
                checkpoint = details["checkpoint"]
                candidate_id = item["candidate_id"]
                if (run["model"] != model or run["status"] != "completed"
                        or saved.values["benchmark_candidate_id"] != candidate_id
                        or saved.values["forecast_horizon_hours"] != horizon
                        or details["evaluation_contract"]["dataset_fingerprint"] != historical_digest):
                    raise ValueError(f"Historical benchmark identity is inconsistent: {run_manifest}")
                fingerprint = stable_signature(training_identity_payload(saved, historical_digest))
                if fingerprint != checkpoint["fingerprint"]:
                    raise ValueError(f"Historical fingerprint does not reproduce: {run_manifest}")
                stage = checkpoint["stage"]
                suffix = "json" if model == "xgboost" else "pt"
                transient_stage = checkpoint_root / model / fingerprint / f"{stage}.{suffix}"
                variants = [
                    current_comparisons(config, saved, historical_digest, current_dataset_digest)
                    for name, config in candidates
                    if config.model == model and (name == candidate_id or name.startswith(candidate_id + "_future_"))
                ]
                if not variants:
                    raise ValueError(f"Current benchmark has no corresponding candidate: {horizon}h {model} {candidate_id}")
                entries.append({
                    "horizon_hours": horizon,
                    "model": model,
                    "historical_candidate_id": candidate_id,
                    "historical_resolved_config": config_path.relative_to(benchmark).as_posix(),
                    "historical_checkpoint_fingerprint": fingerprint,
                    "historical_fingerprint_matches": True,
                    "historical_fit_reported_resumed": checkpoint.get("resumed"),
                    "historical_selected_parameters": item["parameters"],
                    "historical_temporal_split": details["temporal_split"],
                    "historical_final_fit_stage": stage,
                    "original_checkpoint_payload_present": transient_stage.is_file(),
                    "fitted_artifact": fitted_evidence(
                        benchmark, saved, details, item["parameters"], fingerprint, Path(temporary)
                    ),
                    "optuna_evidence": historical_trial_evidence(
                        benchmark, details, study_identity(saved, fingerprint), db_studies
                    ),
                    "current_variants": variants,
                })
    return {
        "contract": "solar-benchmark-model-reuse-audit.v1",
        "benchmark_id": benchmark.name,
        "historical_manifest_sha256": sha256_file(manifest_path),
        "historical_code_revision": provenance["code_revision"],
        "historical_dataset_recorded_fingerprint": historical_digest,
        "historical_dataset_bytes_reverified": current_dataset_digest == historical_digest,
        "current_dataset": {
            "configured_path": str(source),
            "status": "fingerprinted" if current_dataset_digest else "missing",
            "fingerprint": current_dataset_digest,
        },
        "current_experiment": str(experiment),
        "current_experiment_sha256": sha256_file(ROOT / experiment),
        "historical_fingerprints_matched": len(entries),
        "current_config_fingerprints_matched_with_historical_dataset_digest_and_path": sum(
            any(v["same_archived_dataset_digest_and_path_matches"] for v in entry["current_variants"])
            for entry in entries
        ),
        "historical_optuna_sqlite_present": study_db.is_file(),
        "historical_original_checkpoint_payloads_present": sum(entry["original_checkpoint_payload_present"] for entry in entries),
        "fitted_cpu_artifacts_loaded": sum(entry["fitted_artifact"]["cpu_loaded"] for entry in entries),
        "historical_exported_trials": sum(entry["optuna_evidence"]["exported_finished_trials"] for entry in entries),
        "historical_training_resumed_by_audit": False,
        "models_fitted_by_audit": False,
        "comparisons": entries,
    }


def local_evidence_path(value: str, project_root: Path, relative_to: Path) -> Path | None:
    """Resolve a manifest path only inside the explicitly selected evidence root."""
    path = Path(value)
    if PureWindowsPath(value).is_absolute() and not path.is_absolute():
        return None
    candidate = (relative_to / path).resolve()
    return candidate if candidate.is_relative_to(project_root) else None


def inventory_frozen_run(benchmark: Path, project_root: Path) -> dict:
    """Inventory local frozen-run state without loading a training checkpoint."""
    benchmark, project_root = benchmark.resolve(), project_root.resolve()
    manifest_path = benchmark / "manifest.json"
    report = {
        "contract": "solar-frozen-benchmark-reuse-inventory.v1",
        "benchmark": str(benchmark),
        "project_root": str(project_root),
        "manifest_present": manifest_path.is_file(),
        "standard_checkpoint_root_present": (project_root / "artifacts/checkpoints").is_dir(),
        "standard_optuna_db_present": (project_root / "artifacts/optimization/solar_models.db").is_file(),
        "models": [],
        "sqlite_studies": {},
        "checkpoint_payloads_loaded": False,
        "resume_executed": False,
    }
    if not report["manifest_present"]:
        return report
    if not benchmark.is_relative_to(project_root):
        raise ValueError("Frozen benchmark must be inside the selected project directory")
    manifest = read_json(manifest_path)
    report["manifest_sha256"] = sha256_file(manifest_path)
    report["recorded_dataset_fingerprint"] = manifest.get("provenance", {}).get("dataset_fingerprint")
    digests = {}
    gold_errors = {}
    databases = {project_root / "artifacts/optimization/solar_models.db", *benchmark.rglob("optimization.db")}
    for config_path in sorted(benchmark.rglob("resolved_config.json")):
        config_path = artifact_path(benchmark, str(config_path))
        run_directory = config_path.parent
        entry_manifest = run_directory / "manifest.json"
        if not entry_manifest.is_file():
            continue
        entry = read_json(entry_manifest)
        if entry.get("model") not in {"xgboost", "cnn_bilstm"}:
            continue
        details = entry.get("details", {})
        values = read_json(config_path)
        config = ModelJobConfig(entry["model"], BENCHMARK_PROFILE, values, config_path)
        source = local_evidence_path(values["input_dataset"], project_root, project_root)
        gold_status = "missing"
        if source is not None and source.exists() and source not in digests and source not in gold_errors:
            try:
                digests[source] = dataset_signature(source)
            except (OSError, ValueError) as error:
                gold_errors[source] = str(error)
        if source in gold_errors:
            gold_status = "unreadable"
        elif source in digests:
            gold_status = "fingerprinted"
        digest = digests.get(source)
        fingerprint = stable_signature(training_identity_payload(config, digest)) if digest else None
        checkpoint = details.get("checkpoint") or {}
        stored_fingerprint = checkpoint.get("fingerprint")
        stage = checkpoint.get("stage")
        candidate_directories = [run_directory / ".checkpoints", benchmark / ".checkpoints"]
        if checkpoint.get("directory"):
            candidate = local_evidence_path(checkpoint["directory"], project_root, project_root)
            if candidate is not None:
                candidate_directories.append(candidate)
        if isinstance(stored_fingerprint, str) and re.fullmatch(r"[0-9a-f]{64}", stored_fingerprint):
            configured_root = CheckpointSettings.from_config(config).root
            root = project_root / configured_root
            candidate_directories.extend([root / config.model / stored_fingerprint,
                                          run_directory / ".checkpoints" / config.model / stored_fingerprint,
                                          benchmark / ".checkpoints" / config.model / stored_fingerprint])
        for ref, relative_to in ((values.get("optimizer", {}).get("storage_path"), project_root),
                                 ("optimization.db", run_directory),
                                 ("optimization.db", benchmark)):
            if ref:
                db = local_evidence_path(ref, project_root, relative_to)
                if db is not None:
                    databases.add(db)
        optimizer = details.get("optimizer") or {}
        study = study_identity(config, fingerprint) if fingerprint else None
        reported_study = None
        summary_reference = optimizer.get("summary_path")
        if summary_reference:
            summary_path = local_evidence_path(summary_reference, project_root, run_directory)
            summary = read_json(summary_path) if summary_path is not None and summary_path.is_file() else {}
            reported_study = summary.get("study_name")
            if summary.get("storage_path"):
                db = local_evidence_path(summary["storage_path"], project_root, project_root)
                if db is not None:
                    databases.add(db)
        paths = []
        if isinstance(stage, str) and re.fullmatch(r"[A-Za-z0-9_][A-Za-z0-9_.-]*", stage):
            suffix = "json" if config.model == "xgboost" else "pt"
            for directory in sorted(set(candidate_directories)):
                payload = (directory / f"{stage}.{suffix}").resolve()
                if not payload.is_relative_to(project_root):
                    continue
                metadata = payload.with_suffix(payload.suffix + ".meta.json")
                sidecar = read_json(metadata) if metadata.resolve().is_relative_to(project_root) and metadata.is_file() else {}
                signature = sidecar.get("signature")
                paths.append({
                    "path": str(payload),
                    "payload_present": payload.is_file(),
                    "payload_sha256": sha256_file(payload) if payload.is_file() else None,
                    "sidecar_present": bool(sidecar),
                    "sidecar_signature": signature,
                    "sidecar_signature_matches_stage_prefix": (
                        isinstance(signature, str) and stage == f"final_fit_{signature[:20]}"
                    ),
                    "sidecar_identity_matches": bool(sidecar) and all(
                        sidecar.get(key) == expected for key, expected in (
                            ("checkpoint_contract", CHECKPOINT_CONTRACT),
                            ("model", config.model), ("stage", stage), ("fingerprint", fingerprint),
                        )
                    ),
                })
        report["models"].append({
            "model": config.model, "config": str(config_path), "status": entry.get("status"),
            "horizon_hours": values.get("forecast_horizon_hours"),
            "candidate_id": values.get("benchmark_candidate_id"),
            "configured_gold": str(source) if source else None, "gold_status": gold_status,
            "gold_error": gold_errors.get(source),
            "gold_fingerprint": digest, "training_fingerprint": fingerprint,
            "reported_fingerprint": stored_fingerprint,
            "training_fingerprint_matches": fingerprint is not None and fingerprint == stored_fingerprint,
            "stage": stage, "stage_inventory": paths,
            "expected_optuna_study": study, "reported_optuna_study": reported_study,
            "optuna_study_name_matches": study is not None and reported_study == study,
        })
    for path in sorted(databases):
        if not path.resolve().is_relative_to(project_root):
            continue
        try:
            studies = read_only_studies(path)
        except (ValueError, sqlite3.Error) as error:
            report["sqlite_studies"][str(path)] = {"status": "unverified", "reason": str(error)}
        else:
            report["sqlite_studies"][str(path)] = {"present": path.is_file(), "studies": studies}
    for entry in report["models"]:
        study = entry["expected_optuna_study"]
        entry["sqlite_exact_study_locations"] = [
            name for name, db in report["sqlite_studies"].items() if study in db.get("studies", {})
        ]
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--benchmark", type=Path, required=True)
    parser.add_argument("--experiment", type=Path, default=Path("config/experiments/optimized.json"))
    parser.add_argument("--study-db", type=Path)
    parser.add_argument("--checkpoint-root", type=Path)
    parser.add_argument("--inventory", action="store_true", help="Read a frozen local run's Gold and state inventory")
    parser.add_argument("--project-root", type=Path, default=ROOT, help="Evidence root for --inventory")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    archive = args.benchmark.resolve().parents[1]
    if args.inventory:
        report = inventory_frozen_run(args.benchmark, args.project_root)
    else:
        report = audit(
            args.benchmark, args.experiment,
            args.study_db or archive / "optimization/solar_models.db",
            args.checkpoint_root or archive / CheckpointSettings().root.name,
        )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("x", encoding="utf-8") as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
        stream.write("\n")


if __name__ == "__main__":
    main()
