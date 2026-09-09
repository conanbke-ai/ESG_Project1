"""Notification commands: CLI input translation and command dispatch."""
from __future__ import annotations

import argparse
from pathlib import Path
from solar_forecast.config_loader import PROJECT_ROOT


def handle_notify_anomalies_command(args: argparse.Namespace) -> None:
    from solar_forecast.jobs.notification_dispatch_job import (
        NotificationDispatchJob,
        NotificationDispatchJobConfig,
    )

    result = NotificationDispatchJob(project_root=PROJECT_ROOT).run(
        NotificationDispatchJobConfig(
            events=Path(args.events),
            manifest=Path(args.manifest) if args.manifest else None,
            routes=Path(args.routes),
            outbox=Path(args.outbox) if args.outbox else None,
            template_contract=args.template_contract,
            live=args.live,
        )
    )
    print(
        f"Notification {result.mode}: {result.event_count} events verified, "
        f"{result.queued} queued, {result.duplicates} duplicates"
    )
    print(
        f"Provider accepted: {result.accepted}, dry-run suppressed: "
        f"{result.suppressed}, retry pending: {result.retry_required}, "
        f"fallback scheduled: {result.fallback_scheduled}, "
        f"dead-lettered: {result.dead_lettered}, "
        f"provider outcome unknown: {result.in_doubt}"
    )
    print("Provider acceptance is not the same as handset delivery confirmation.")
