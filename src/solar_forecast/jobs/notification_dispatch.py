from __future__ import annotations

from dataclasses import dataclass, replace
from pathlib import Path
from typing import Callable

from solar_forecast.anomalies import verify_operational_event_batch
from solar_forecast.infrastructure.local_env import load_local_env
from solar_forecast.notifications import (
    AnomalyAlert,
    NotificationDispatcher,
    NotificationOutbox,
    NotificationService,
    NotificationSettings,
    RuntimeRouteDirectory,
    SolapiSettings,
    build_solapi_dispatcher,
)
from solar_forecast.settings import PROJECT_ROOT


@dataclass(frozen=True)
class NotificationDispatchJobConfig:
    events: Path
    routes: Path
    manifest: Path | None = None
    outbox: Path | None = None
    template_contract: str = "solar-anomaly-v1"
    live: bool = False


@dataclass(frozen=True)
class NotificationDispatchJobResult:
    mode: str
    event_count: int
    queued: int
    duplicates: int
    accepted: int
    suppressed: int
    retry_required: int
    fallback_scheduled: int
    dead_lettered: int
    in_doubt: int


class NotificationDispatchJob:
    """Validate operational anomaly events and dispatch them through an outbox.

    This is intentionally a job boundary rather than a general microservice.  The
    CLI, a future queue consumer, or a container entrypoint can call the same
    object without mixing external side effects into dashboard/model code.
    """

    def __init__(
        self,
        *,
        project_root: Path = PROJECT_ROOT,
        env_loader: Callable[[], None] = load_local_env,
    ):
        self.project_root = Path(project_root)
        self.env_loader = env_loader

    def run(self, config: NotificationDispatchJobConfig) -> NotificationDispatchJobResult:
        self.env_loader()
        settings = NotificationSettings.from_env().confirm_live(confirmed=config.live)
        if config.live and not settings.external_delivery_allowed:
            raise SystemExit(
                "Live notifications are blocked. Set SOLAR_NOTIFY_ENABLED=true and "
                "SOLAR_NOTIFY_DRY_RUN=false, then pass --live explicitly."
            )

        events_path = self._project_path(config.events)
        manifest_path = self._project_path(config.manifest) if config.manifest else None
        routes_path = self._project_path(config.routes)
        batch = verify_operational_event_batch(events_path, manifest_path)

        if config.outbox and not config.live:
            raise SystemExit(
                "Dry-run preview does not accept --outbox. It always uses the "
                "isolated artifacts/notifications/dry_run.sqlite3 database."
            )
        if config.live:
            outbox_path = self._project_path(config.outbox or settings.outbox_path)
        else:
            # A preview must not consume the live idempotency key and suppress a
            # later, explicitly approved delivery of the same operational event.
            outbox_path = self.project_root / "artifacts/notifications/dry_run.sqlite3"
        settings = replace(settings, outbox_path=outbox_path)
        routes = RuntimeRouteDirectory.from_file(
            routes_path,
            require_contacts=settings.external_delivery_allowed,
        )
        solapi = SolapiSettings.from_env() if settings.external_delivery_allowed else None

        def validated_alerts():
            # Re-open the verified JSONL for each pass so semantic validation is
            # fail-closed without retaining a large event batch in memory.
            for source in batch.records():
                record = dict(source)
                record.setdefault("run_id", batch.run_id)
                record.setdefault("detector_version", batch.detector_version)
                alert = AnomalyAlert.from_operational_event(record)
                yield alert, routes.routes_for(
                    alert.route_key,
                    alert.audience,
                    alert.plant_id,
                )

        # Validate every event and route before the outbox is created or mutated.
        for _alert, _routes in validated_alerts():
            pass

        outbox = NotificationOutbox(outbox_path)
        service = NotificationService(outbox, template_id=config.template_contract)

        enqueued = 0
        duplicates = 0
        for alert, alert_routes in validated_alerts():
            for result in service.enqueue_anomaly(alert, alert_routes):
                if result.created:
                    enqueued += 1
                else:
                    duplicates += 1

        if settings.external_delivery_allowed:
            dispatcher = build_solapi_dispatcher(
                settings,
                outbox,
                routes.contacts,
                solapi,
            )
        else:
            dispatcher = NotificationDispatcher(settings, outbox, routes.contacts, {})

        totals = {
            "candidates": 0,
            "accepted": 0,
            "suppressed": 0,
            "retry_required": 0,
            "fallback_scheduled": 0,
            "dead_lettered": 0,
            "in_doubt": 0,
        }
        while True:
            summary = dispatcher.run_once()
            for key in totals:
                totals[key] += int(getattr(summary, key))
            if summary.candidates == 0:
                break

        return NotificationDispatchJobResult(
            mode="LIVE" if settings.external_delivery_allowed else "DRY-RUN",
            event_count=batch.event_count,
            queued=enqueued,
            duplicates=duplicates,
            accepted=totals["accepted"],
            suppressed=totals["suppressed"],
            retry_required=totals["retry_required"],
            fallback_scheduled=totals["fallback_scheduled"],
            dead_lettered=totals["dead_lettered"],
            in_doubt=totals["in_doubt"],
        )

    def _project_path(self, value: str | Path) -> Path:
        path = Path(value)
        return path if path.is_absolute() else self.project_root / path


__all__ = [
    "NotificationDispatchJob",
    "NotificationDispatchJobConfig",
    "NotificationDispatchJobResult",
]
