"""Official-source data collectors."""
from importlib import import_module

_EXPORTS = {'CollectionConfig': ('solar_forecast.collectors.collection_config', 'CollectionConfig'), 'CollectionService': ('solar_forecast.collectors.collection_service', 'CollectionService'), 'collect_all': ('solar_forecast.collectors.collection_service', 'collect_all'), 'CollectedGenerationAdmissionResult': ('solar_forecast.datasets.collector_admission', 'CollectedGenerationAdmissionResult'), 'CollectedGenerationAdmissionService': ('solar_forecast.datasets.collector_admission', 'CollectedGenerationAdmissionService'), 'DailyWideGenerationNormalizer': ('solar_forecast.collectors.generation_normalizers', 'DailyWideGenerationNormalizer'), 'EwpTrainingNormalizer': ('solar_forecast.collectors.generation_normalizers', 'EwpTrainingNormalizer'), 'KrcYeongamGenerationNormalizer': ('solar_forecast.collectors.generation_normalizers', 'KrcYeongamGenerationNormalizer'), 'KoenGenerationNormalizer': ('solar_forecast.collectors.generation_normalizers', 'KoenGenerationNormalizer'), 'HistoricalGenerationStandardizationService': ('solar_forecast.datasets.archive_standardizer', 'HistoricalGenerationStandardizationService'), 'KrcYeongamCandidateIntakeService': ('solar_forecast.datasets.krc_candidate_intake', 'KrcYeongamCandidateIntakeService'), 'KomipoRenewableCollector': ('solar_forecast.collectors.komipo_api', 'KomipoRenewableCollector')}
__all__ = list(_EXPORTS)

def __getattr__(name: str):
    """Load a public symbol only when its owning feature is requested."""
    if name not in _EXPORTS:
        raise AttributeError(name)
    module, symbol = _EXPORTS[name]
    value = getattr(import_module(module), symbol)
    globals()[name] = value
    return value
