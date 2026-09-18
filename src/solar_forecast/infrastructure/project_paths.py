"""Default storage locations; explicit CLI/configuration paths take precedence.

Retained provider archives keep their existing paths and hashes. New experiment
artifacts share artifacts/; source code never doubles as a training output folder.
"""
from pathlib import Path

GENERATION_ARCHIVE_ROOT = Path("file/solar_data_file")
WEATHER_ARCHIVE_ROOT = Path("file/KMA_data_file")
LEGACY_MERGED_ROOT = Path("file/merge_data")
LEGACY_MERGED_SOURCE = LEGACY_MERGED_ROOT / "val.csv"
BRONZE_ROOT = Path("file/raw")
SILVER_ROOT = Path("file/standardized")
COLLECTOR_SILVER_ROOT = SILVER_ROOT / "downloads"
GOLD_DATASET_PATH = SILVER_ROOT / "model_ready.csv.gz"
GOLD_PARTITIONS_ROOT = SILVER_ROOT / "model_ready_parts"
PLANT_REGISTRY_PATH = SILVER_ROOT / "plant_registry.csv"
PLANT_QUALITY_PATH = SILVER_ROOT / "plant_quality_report.csv"
PIPELINE_RUNS_ROOT = Path("artifacts/pipeline")
HYBRID_EXPERIMENT_ROOT = Path("artifacts/experiments/hybrid")
FEATURE_EVALUATION_ROOT = Path("artifacts/evaluation/features")
VERIFICATION_ROOT = Path("artifacts/verification/e2e")
DASHBOARD_ROOT = Path("dashboard")
