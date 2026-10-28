"""reporting public API."""
from importlib import import_module

_EXPORTS = {'DashboardBuildResult': ('solar_forecast.reporting.dashboard_builder', 'DashboardBuildResult'), 'DashboardBuilder': ('solar_forecast.reporting.dashboard_builder', 'DashboardBuilder'), 'ModelAnalyticsService': ('solar_forecast.reporting.model_analytics', 'ModelAnalyticsService'), 'NationalInventoryService': ('solar_forecast.reporting.national_solar_inventory', 'NationalInventoryService'), 'build_national_inventory': ('solar_forecast.reporting.national_solar_inventory', 'build_national_inventory'), 'ProvinceBoundaryError': ('solar_forecast.reporting.sgis_boundaries', 'ProvinceBoundaryError'), 'SgisBoundarySource': ('solar_forecast.reporting.sgis_boundaries', 'SgisBoundarySource'), 'SgisProvinceBoundaryConverter': ('solar_forecast.reporting.sgis_boundaries', 'SgisProvinceBoundaryConverter'), 'validate_province_boundaries': ('solar_forecast.reporting.sgis_boundaries', 'validate_province_boundaries'), 'verify_sgis_source_bundle': ('solar_forecast.reporting.sgis_boundaries', 'verify_sgis_source_bundle')}
__all__ = list(_EXPORTS)

def __getattr__(name: str):
    """Load a public symbol only when its owning feature is requested."""
    if name not in _EXPORTS:
        raise AttributeError(name)
    module, symbol = _EXPORTS[name]
    value = getattr(import_module(module), symbol)
    globals()[name] = value
    return value
