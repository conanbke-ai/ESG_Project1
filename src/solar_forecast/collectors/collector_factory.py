"""Construct only the selected provider adapter and its optional dependencies."""
from __future__ import annotations

import os

from solar_forecast.collectors.collection_config import CollectionConfig, load_source_catalog
from solar_forecast.collectors.contracts import Collector


def build_collector(source: str, config: CollectionConfig) -> Collector:
    """Resolve one source, selecting ASOS API mode after local environment loading."""
    if source == "kma":
        from solar_forecast.collectors.kma_api import KmaAsosApiCollector
        from solar_forecast.collectors.kma_api_client import ASOS_SERVICE_KEY_ENV
        if config.kma_mode == "api" or (config.kma_mode == "auto" and os.getenv(ASOS_SERVICE_KEY_ENV, "").strip()):
            return KmaAsosApiCollector(config)
        from solar_forecast.collectors.kma_browser import KmaAsosBrowserCollector
        return KmaAsosBrowserCollector(config)
    if source == "komipo":
        from solar_forecast.collectors.komipo_api import KomipoRenewableCollector
        return KomipoRenewableCollector(config)
    from solar_forecast.collectors.generation_collectors import (
        DataGoDatasetSpec, DataGoFileCollector, EwpAttachmentSpec,
        EwpTrainingDataCollector, KoenHomepageCollector,
    )
    if source == "koen":
        return KoenHomepageCollector(config)
    metadata = load_source_catalog()[source]
    if source == "ewp":
        return EwpTrainingDataCollector(config, EwpAttachmentSpec(
            detail_url=metadata["dataset_page"], download_url=metadata["download_url"],
            attachment_id=metadata["attachment_id"], order_number=metadata["order_number"],
            page_code=metadata["page_code"], organization=metadata["organization"],
            detail_name=metadata["detail_name"],
        ))
    if source in {"kospo", "iwest"}:
        from solar_forecast.collectors.generation_normalizers import DailyWideGenerationNormalizer, IWEST_WIDE_SCHEMA, KOSPO_WIDE_SCHEMA
        spec = DataGoDatasetSpec(source, metadata["dataset_id"], metadata["detail_id"], metadata["organization"], metadata["detail_name"])
        schema = KOSPO_WIDE_SCHEMA if source == "kospo" else IWEST_WIDE_SCHEMA
        return DataGoFileCollector(spec, config, DailyWideGenerationNormalizer(schema))
    raise ValueError(f"Unsupported collector: {source}")
