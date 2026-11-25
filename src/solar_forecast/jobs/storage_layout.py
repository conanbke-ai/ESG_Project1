"""Read-only ownership projection of existing storage paths and job contracts.

This catalog describes default locations, not permission enforcement, dataset
admission, integrity verification or a filesystem migration. It never reads data
records or creates, moves, deletes or rewrites an artifact.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass
from enum import StrEnum
import os
from pathlib import Path
from typing import Any

from solar_forecast.infrastructure import project_paths
from solar_forecast.jobs.contracts import get_job_contract, list_job_contracts


STORAGE_LAYOUT_CONTRACT = "solar-storage-layout.v1"


class StorageLayer(StrEnum):
    """Logical roles; they do not rename existing physical directories."""

    BRONZE = "bronze"
    SILVER = "silver"
    GOLD = "gold"
    ARTIFACTS = "artifacts"
    REPORTING = "reporting"
    REFERENCE = "reference"
    LEGACY = "legacy"
    SOURCE = "source"
    CONFIG = "config"


@dataclass(frozen=True)
class StorageLocation:
    """One declared path owner; file roots match exactly, directories by ancestry."""

    storage_id: str
    path: str
    kind: str
    layer: StorageLayer
    owner: str
    preservation_policy: str

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


def _location(storage_id: str, path: Path | str, kind: str, layer: StorageLayer,
              owner: str, policy: str) -> StorageLocation:
    return StorageLocation(storage_id, Path(path).as_posix(), kind, layer, owner, policy)


def list_storage_locations() -> tuple[StorageLocation, ...]:
    """Return existing default paths without discovering files or loading records."""

    layer = StorageLayer
    gold_manifest = get_job_contract("prepare-data").outputs[0].path_pattern
    return (
        _location("generation_archive", project_paths.GENERATION_ARCHIVE_ROOT, "directory", layer.BRONZE, "collectors", "preserve retained source bytes/revisions"),
        _location("weather_archive", project_paths.WEATHER_ARCHIVE_ROOT, "directory", layer.BRONZE, "collectors", "existing collector merge/revision contract only"),
        _location("bronze", project_paths.BRONZE_ROOT, "directory", layer.BRONZE, "collectors", "preserve received source and acquisition lineage"),
        _location("silver", project_paths.SILVER_ROOT, "directory", layer.SILVER, "preprocessing", "explicit preparation job only; not model admission"),
        _location("collector_silver", project_paths.COLLECTOR_SILVER_ROOT, "directory", layer.SILVER, "collectors", "admission required before model use"),
        _location("plant_registry", project_paths.PLANT_REGISTRY_PATH, "file", layer.SILVER, "datasets", "preserve reviewed identity/weather mapping contract"),
        _location("plant_quality", project_paths.PLANT_QUALITY_PATH, "file", layer.SILVER, "quality", "quality evidence is not an automatic training approval"),
        _location("gold_dataset", project_paths.GOLD_DATASET_PATH, "file", layer.GOLD, "datasets", "preserve columns/units/feature order and lineage"),
        _location("gold_partitions", project_paths.GOLD_PARTITIONS_ROOT, "directory", layer.GOLD, "datasets", "reuse only under existing dataset/config fingerprint"),
        _location("gold_manifest", gold_manifest, "file", layer.GOLD, "datasets", "existing model-ready manifest contract remains authoritative"),
        _location("legacy_merged", project_paths.LEGACY_MERGED_ROOT, "directory", layer.LEGACY, "datasets", "compatibility only; no implicit newest-data or Gold promotion"),
        _location("reference_data", "data", "directory", layer.REFERENCE, "reporting", "mixed historical caches/references; review before reuse"),
        _location("artifacts", "artifacts", "directory", layer.ARTIFACTS, "jobs", "retain run provenance; catalog grants no write/delete authority"),
        _location("model_runs", "artifacts/models", "directory", layer.ARTIFACTS, "models", "preserve completed packages/metrics and manifest identity"),
        _location("checkpoints", "artifacts/checkpoints", "directory", layer.ARTIFACTS, "models.shared", "existing checkpoint fingerprint/resume lifecycle only"),
        _location("optuna", "artifacts/optimization", "directory", layer.ARTIFACTS, "models.shared", "existing study identity and official stale recovery only"),
        _location("pipeline_runs", project_paths.PIPELINE_RUNS_ROOT, "directory", layer.ARTIFACTS, "pipeline", "preserve run IDs and completed results"),
        _location("hybrid_experiments", project_paths.HYBRID_EXPERIMENT_ROOT, "directory", layer.ARTIFACTS, "models.hybrid", "preserve validation/calibration/test role separation"),
        _location("feature_evaluation", project_paths.FEATURE_EVALUATION_ROOT, "directory", layer.ARTIFACTS, "evaluation", "report-only results do not authorize model reselection"),
        _location("verification", project_paths.VERIFICATION_ROOT, "directory", layer.ARTIFACTS, "jobs", "verification report does not imply formal model performance PASS"),
        _location("dashboard", project_paths.DASHBOARD_ROOT, "directory", layer.REPORTING, "reporting", "published projection; never a source-of-truth training dataset"),
        _location("map", "map", "directory", layer.REFERENCE, "reporting", "preserve source/license/hash and legacy references"),
        _location("map_redirects", "map/html", "directory", layer.REPORTING, "reporting", "retain old URL compatibility"),
        _location("source", "src", "directory", layer.SOURCE, "package", "no runtime artifacts inside source"),
        _location("config", "config", "directory", layer.CONFIG, "configuration", "preserve persisted model/feature settings and secret exclusions"),
        *(_location("redirect_" + Path(name).stem, name, "file", layer.REPORTING, "reporting", "retain old URL compatibility")
          for name in ("solar_dashboard.html", "forecast.html", "model_analysis.html", "plant_region_report_perm.html")),
    )


def _find_location(relative_path: Path, locations: tuple[StorageLocation, ...]) -> StorageLocation | None:
    matches = []
    for location in locations:
        declared = Path(location.path)
        if location.kind == "file":
            if relative_path != declared and relative_path.is_relative_to(declared):
                raise ValueError("A declared file cannot be used as a parent directory")
            matches_location = relative_path == declared
        else:
            matches_location = relative_path.is_relative_to(declared)
        if matches_location:
            matches.append(location)
    return max(matches, key=lambda item: len(Path(item.path).parts), default=None)


def classify_storage_path(project_root: Path, requested_path: str | Path) -> dict[str, Any]:
    """Describe one contained path, resolving symlinks but never reading its content.

    Unknown paths remain unclassified. Explicit CLI/config overrides are not moved,
    rejected as model inputs or adopted by this descriptive default-path catalog.
    """

    if "://" in str(requested_path):
        raise ValueError("A project filesystem path is required, not a URL")
    root = Path(project_root).resolve()
    path = Path(requested_path)
    # Reject lexical escapes (including external drive/UNC paths on Windows)
    # before asking the filesystem to resolve an untrusted candidate.
    contained = Path(os.path.abspath(path if path.is_absolute() else root / path))
    if not contained.is_relative_to(root):
        raise ValueError("Path is outside the project root")
    resolved = contained.resolve()
    if not resolved.is_relative_to(root):
        raise ValueError("Path is outside the project root")
    relative = resolved.relative_to(root)
    location = _find_location(relative, list_storage_locations())
    return {
        "path": relative.as_posix(),
        "classification": "declared_default_location" if location else "unclassified_review_required",
        "location": location.as_dict() if location else None,
        "data_integrity_checked": False,
        "model_eligibility_checked": False,
    }


def build_storage_layout_catalog(project_root: Path, requested_path: str | Path | None = None) -> dict[str, Any]:
    """Build a read-only CLI projection, reusing versioned JobContract definitions.

    No filesystem enumeration, data parsing, hashing, network calls, lock/state
    mutation, collection, model import, training or notification dispatch occurs.
    """

    locations = list_storage_locations()
    references = []
    for job in list_job_contracts():
        for direction, artifacts in (("input", job.inputs), ("output", job.outputs)):
            for artifact in artifacts:
                prefix = artifact.path_pattern.split("<", 1)[0].rstrip("/")
                location = _find_location(Path(prefix), locations)
                references.append({
                    "job_id": job.job_id, "direction": direction, "artifact": artifact.name,
                    "contract": artifact.contract, "path_pattern": artifact.path_pattern,
                    "storage_id": location.storage_id if location else None,
                    "required_for_next_job": artifact.required_for_next_job,
                })
    return {
        "contract": STORAGE_LAYOUT_CONTRACT, "schema_version": 1,
        "scope": "declared_default_path_ownership_only", "read_only": True,
        "automatic_move_allowed": False, "automatic_delete_allowed": False,
        "write_protection_enforced": False,
        "locations": [location.as_dict() for location in locations],
        "job_artifact_references": references,
        "query": classify_storage_path(project_root, requested_path) if requested_path is not None else None,
    }
