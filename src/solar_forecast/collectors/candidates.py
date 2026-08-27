from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from solar_forecast.evaluation.temporal import TemporalSplitConfig, TemporalSplitter
from solar_forecast.artifacts.manifest import sha256_file, write_json_atomic

from .normalization import (
    GENERATION_COLUMNS,
    GENERATION_CONTRACT_VERSION,
    KrcYeongamGenerationNormalizer,
    read_csv_with_fallback,
)


@dataclass(frozen=True)
class CandidateAcceptancePolicy:
    minimum_days: int = 3 * 365
    minimum_hourly_coverage: float = 0.95
    maximum_gap_hours: int = 24
    gap_hours: int = 168
    maximum_hourly_capacity_factor: float = 1.05


@dataclass(frozen=True)
class CandidateSourceFile:
    source: str
    bytes: int
    sha256: str
    daily_rows: int
    status: str
    reason: str | None


@dataclass(frozen=True)
class CandidatePlantProfile:
    plant_id: str
    start: str
    end: str
    observed_hours: int
    expected_hours: int
    hourly_coverage: float
    maximum_gap_hours: int
    capacity_coverage: float
    maximum_hourly_capacity_factor: float | None
    negative_generation: int
    split_rows: dict[str, int]
    generation_gate_passed: bool
    issues: tuple[str, ...]


@dataclass(frozen=True)
class CandidateIntakeResult:
    manifest_path: Path
    dataset_path: Path
    rows: int
    plants: int
    status: str
    source_files: tuple[CandidateSourceFile, ...]
    profiles: tuple[CandidatePlantProfile, ...]


class KrcYeongamCandidateIntakeService:
    """Stage, standardize, and gate KRC Yeongam files before model admission.

    Generation quality and weather join readiness are deliberately separate.
    Passing this service does not silently append a plant to model training;
    an explicit ASOS mapping and source-unit review must still be approved.
    """

    def __init__(
        self,
        source_dir: Path,
        output_dir: Path,
        policy: CandidateAcceptancePolicy | None = None,
    ):
        self.source_dir = Path(source_dir)
        self.output_dir = Path(output_dir)
        self.policy = policy or CandidateAcceptancePolicy()
        self.normalizer = KrcYeongamGenerationNormalizer()

    def run(self) -> CandidateIntakeResult:
        sources = sorted(self.source_dir.glob("*.csv"))
        if not sources:
            raise FileNotFoundError(f"No candidate CSV files found: {self.source_dir}")

        accepted_parts: list[pd.DataFrame] = []
        source_results: list[CandidateSourceFile] = []
        for source in sources:
            raw = read_csv_with_fallback(source)
            try:
                normalized = self.normalizer.read(source)
                accepted_parts.append(normalized)
                status, reason = "accepted_for_generation_audit", None
            except ValueError as exc:
                status, reason = "quarantined", str(exc)
            source_results.append(
                CandidateSourceFile(
                    source=str(source),
                    bytes=source.stat().st_size,
                    sha256=sha256_file(source),
                    daily_rows=len(raw),
                    status=status,
                    reason=reason,
                )
            )
        if not accepted_parts:
            raise ValueError("Every KRC candidate file was quarantined")

        generation = pd.concat(accepted_parts, ignore_index=True)
        generation["timestamp"] = pd.to_datetime(generation["timestamp"], errors="coerce")
        generation = generation.sort_values(["timestamp", "plant_id"], kind="stable")
        duplicate_keys = int(generation.duplicated(["timestamp", "plant_id"]).sum())
        if duplicate_keys:
            raise ValueError(f"Candidate files overlap on {duplicate_keys} hourly keys")

        splitter = TemporalSplitter(
            TemporalSplitConfig(
                validation_fraction=0.15,
                calibration_fraction=0.10,
                test_fraction=0.15,
                gap_hours=self.policy.gap_hours,
            )
        )
        splits = splitter.split_frame(generation)
        split_frames = {
            "train": splits.train,
            "validation": splits.validation,
            "calibration": splits.calibration,
            "test": splits.test,
        }
        profiles = tuple(
            self._profile_plant(plant_id, group, split_frames)
            for plant_id, group in generation.groupby("plant_id", sort=True)
        )

        self.output_dir.mkdir(parents=True, exist_ok=True)
        partition_dir = self.output_dir / "generation"
        partition_dir.mkdir(parents=True, exist_ok=True)
        partition_paths: list[Path] = []
        for year, part in generation.groupby(generation["timestamp"].dt.year, sort=True):
            destination = partition_dir / f"krc_yeongam_{int(year)}.csv.gz"
            self._write_csv_atomic(part, destination)
            partition_paths.append(destination)

        dataset_path = self.output_dir / "candidate_generation.csv.gz"
        self._write_csv_atomic(generation, dataset_path)
        generation_ready = all(profile.generation_gate_passed for profile in profiles)
        status = "generation_ready_for_registry" if generation_ready else "generation_quality_review_required"
        manifest = {
            "created_at": datetime.now().isoformat(),
            "source_catalog": "https://www.data.go.kr/dataset/15005796/fileData.do?lang=ko",
            "contract_version": GENERATION_CONTRACT_VERSION,
            "contract": GENERATION_COLUMNS,
            "processing_model": "source-file bounded normalization + year-partitioned gzip",
            "source_files": [asdict(item) for item in source_results],
            "partitions": [
                {
                    "path": str(path),
                    "bytes": path.stat().st_size,
                    "sha256": sha256_file(path),
                }
                for path in partition_paths
            ],
            "dataset": str(dataset_path),
            "rows": len(generation),
            "plants": int(generation["plant_id"].nunique()),
            "duplicate_keys": duplicate_keys,
            "negative_generation": int(generation["generation_mwh"].lt(0).sum()),
            "source_unit_interpretation": (
                "portal-declared kW one-hour buckets are energy-equivalent kWh buckets; "
                "converted to MWh by dividing by 1000"
            ),
            "unit_validation": {
                "portal_contract": "kW hourly buckets with a reconciled 24-hour daily total",
                "physical_bound": "every hourly MWh value must be <= 1.05 * official MW capacity",
                "passed": generation_ready,
            },
            "unit_review_required": False,
            "weather_join": {
                "ready": False,
                "reason": "resolved later by the version-controlled reviewed station mapping registry",
            },
            "split_protocol": {
                "boundaries": splits.boundaries.to_dict(),
                "rows": {name: len(frame) for name, frame in split_frames.items()},
                "roles": {
                    "train": "model fitting and preprocessing statistics",
                    "validation": "model/feature/hybrid-gate selection",
                    "calibration": "frozen residual thresholds and interval calibration",
                    "test": "final one-time reporting",
                },
            },
            "profiles": [asdict(profile) for profile in profiles],
            "status": status,
            "training_admission": (
                "delegated to the nationwide plant registry; only reviewed weather mappings are admitted "
                "and feature columns remain identical to every other plant"
            ),
        }
        manifest_path = self.output_dir / "candidate_manifest.json"
        write_json_atomic(manifest_path, manifest)
        return CandidateIntakeResult(
            manifest_path=manifest_path,
            dataset_path=dataset_path,
            rows=len(generation),
            plants=int(generation["plant_id"].nunique()),
            status=status,
            source_files=tuple(source_results),
            profiles=profiles,
        )

    def _profile_plant(
        self,
        plant_id: str,
        frame: pd.DataFrame,
        split_frames: dict[str, pd.DataFrame],
    ) -> CandidatePlantProfile:
        timestamps = frame["timestamp"].drop_duplicates().sort_values()
        start, end = timestamps.iloc[0], timestamps.iloc[-1]
        expected_hours = int((end - start) / pd.Timedelta(hours=1)) + 1
        gaps = timestamps.diff().dropna().div(pd.Timedelta(hours=1)).sub(1)
        maximum_gap = int(gaps.max()) if not gaps.empty else 0
        coverage = len(timestamps) / expected_hours
        days = int((end.normalize() - start.normalize()) / pd.Timedelta(days=1)) + 1
        issues: list[str] = []
        if days < self.policy.minimum_days:
            issues.append(f"history_days<{self.policy.minimum_days}")
        if coverage < self.policy.minimum_hourly_coverage:
            issues.append(f"hourly_coverage<{self.policy.minimum_hourly_coverage}")
        if maximum_gap > self.policy.maximum_gap_hours:
            issues.append(f"maximum_gap_hours>{self.policy.maximum_gap_hours}")
        capacity_coverage = float(frame["capacity_mw"].notna().mean())
        if capacity_coverage < 1:
            issues.append("capacity_missing")
        capacity_factor = frame["generation_mwh"].div(
            pd.to_numeric(frame["capacity_mw"], errors="coerce")
        )
        maximum_capacity_factor = (
            float(capacity_factor.max()) if capacity_factor.notna().any() else None
        )
        if (
            maximum_capacity_factor is not None
            and maximum_capacity_factor > self.policy.maximum_hourly_capacity_factor
        ):
            issues.append(
                f"hourly_capacity_factor>{self.policy.maximum_hourly_capacity_factor}"
            )
        negative = int(frame["generation_mwh"].lt(0).sum())
        if negative:
            issues.append("negative_generation")
        split_rows = {
            name: int(part["plant_id"].eq(plant_id).sum())
            for name, part in split_frames.items()
        }
        if any(value == 0 for value in split_rows.values()):
            issues.append("empty_temporal_partition")
        return CandidatePlantProfile(
            plant_id=str(plant_id),
            start=start.isoformat(),
            end=end.isoformat(),
            observed_hours=len(timestamps),
            expected_hours=expected_hours,
            hourly_coverage=coverage,
            maximum_gap_hours=maximum_gap,
            capacity_coverage=capacity_coverage,
            maximum_hourly_capacity_factor=maximum_capacity_factor,
            negative_generation=negative,
            split_rows=split_rows,
            generation_gate_passed=not issues,
            issues=tuple(issues),
        )

    @staticmethod
    def _write_csv_atomic(frame: pd.DataFrame, destination: Path) -> None:
        temporary = destination.with_name(destination.name + ".tmp")
        frame.to_csv(
            temporary,
            index=False,
            encoding="utf-8-sig",
            compression={"method": "gzip", "compresslevel": 1, "mtime": 1},
        )
        temporary.replace(destination)
