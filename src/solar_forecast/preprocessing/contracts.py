"""Versioned contracts for the Solar plant-data preprocessing pipeline."""
from __future__ import annotations

from enum import StrEnum

PLANT_PREPROCESSING_POLICY_CONTRACT = "solar-plant-preprocessing-policy.v1"
OBSERVED_PREPROCESSING_CONTRACT = "solar-observed-preprocessing.v2"
TRAINING_ELIGIBILITY_CONTRACT = "solar-training-data-eligibility.v2"
TRAINING_ADMISSION_POLICY_CONTRACT = "solar-training-admission-policy.v2"
TRAINING_POPULATION_CONTRACT = "solar-training-population.v2"
ADMITTED_DATASET_CONTRACT = "solar-admitted-training-dataset.v1"

class PreprocessingStage(StrEnum):
    SOURCE_PRESERVATION = "source_preservation"
    STANDARDIZATION = "standardization"
    IDENTITY_AND_WEATHER_MAPPING = "identity_and_weather_mapping"
    GOLD_JOIN = "gold_join"
    QUALITY_ANNOTATION = "quality_annotation"
    LEAKAGE_SAFE_FEATURES = "leakage_safe_features"
    TRAINING_ELIGIBILITY = "training_eligibility"
    TRAINING_ADMISSION = "training_admission"
    ADMITTED_DATASET = "admitted_dataset"
    MODEL_READINESS = "model_readiness"

CANONICAL_STAGE_ORDER = tuple(PreprocessingStage)

__all__ = [
    "PLANT_PREPROCESSING_POLICY_CONTRACT",
    "OBSERVED_PREPROCESSING_CONTRACT",
    "TRAINING_ELIGIBILITY_CONTRACT",
    "TRAINING_ADMISSION_POLICY_CONTRACT",
    "TRAINING_POPULATION_CONTRACT",
    "ADMITTED_DATASET_CONTRACT",
    "PreprocessingStage",
    "CANONICAL_STAGE_ORDER",
]
