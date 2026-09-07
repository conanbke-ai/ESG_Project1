"""Official-source data collectors."""

from .config import CollectionConfig
from .service import CollectionService, collect_all
from .admission import (
    CollectedGenerationAdmissionResult,
    CollectedGenerationAdmissionService,
)
from .normalization import (
    DailyWideGenerationNormalizer,
    EwpTrainingNormalizer,
    KrcYeongamGenerationNormalizer,
    KoenGenerationNormalizer,
)
from .archive import HistoricalGenerationStandardizationService
from .candidates import KrcYeongamCandidateIntakeService
from .openapi import KomipoRenewableCollector

__all__ = [
    "CollectedGenerationAdmissionResult",
    "CollectedGenerationAdmissionService",
    "CollectionConfig",
    "CollectionService",
    "DailyWideGenerationNormalizer",
    "EwpTrainingNormalizer",
    "HistoricalGenerationStandardizationService",
    "KrcYeongamCandidateIntakeService",
    "KrcYeongamGenerationNormalizer",
    "KoenGenerationNormalizer",
    "KomipoRenewableCollector",
    "collect_all",
]
