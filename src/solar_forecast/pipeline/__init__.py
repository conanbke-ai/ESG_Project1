"""Composable forecasting pipeline with lazy framework imports."""

from typing import Any

__all__ = ["ForecastPipeline", "PipelineConfig", "PipelineResult", "run_pipeline"]


def __getattr__(name: str) -> Any:
    if name == "PipelineConfig":
        from .config import PipelineConfig
        return PipelineConfig
    if name in {"ForecastPipeline", "PipelineResult", "run_pipeline"}:
        from .service import ForecastPipeline, PipelineResult, run_pipeline
        return {"ForecastPipeline": ForecastPipeline, "PipelineResult": PipelineResult, "run_pipeline": run_pipeline}[name]
    raise AttributeError(name)
