"""이벤트 enqueue와 재시도·대체 채널·전송 결과 전이 처리."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Mapping, Sequence

from solar_forecast.notifications.notification_config import NotificationSettings, SolapiSettings
from solar_forecast.notifications.contracts import (
    AnomalyAlert,
    NotificationAudience,
    NotificationChannel,
    NotificationEnvelope,
    NotificationStatus,
    RecipientRoute,
)
from solar_forecast.notifications.outbox_repository import EnqueueResult, NotificationOutbox
from solar_forecast.notifications.delivery_provider import (
    AmbiguousNotificationError,
    ContactDirectory,
    NotificationProvider,
    PermanentNotificationError,
    ProviderMessage,
    SolapiProvider,
    TransientNotificationError,
)


@dataclass(frozen=True)
class DispatchSummary:
    candidates: int = 0
    accepted: int = 0
    suppressed: int = 0
    retry_required: int = 0
    fallback_scheduled: int = 0
    dead_lettered: int = 0
    in_doubt: int = 0


class NotificationService:
    """Convert validated operational anomalies into deduplicated outbox work."""

    def __init__(self, outbox: NotificationOutbox, *, template_id: str):
        self.outbox = outbox
        self.template_id = str(template_id).strip()
        if not self.template_id:
            raise ValueError("notification template_id cannot be blank")

    def enqueue_anomaly(
        self,
        alert: AnomalyAlert,
        routes: Sequence[RecipientRoute],
        *,
        now: datetime | None = None,
    ) -> list[EnqueueResult]:
        if alert.scope != "operational":
            raise ValueError("evaluation/test anomalies cannot enter the alert outbox")
        if not routes:
            raise ValueError("at least one recipient route is required")
        results: list[EnqueueResult] = []
        seen: set[str] = set()
        for route in routes:
            if route.route_key != alert.route_key:
                raise ValueError("recipient route_key does not match the anomaly route_key")
            if route.audience != alert.audience:
                raise ValueError("recipient audience does not match the anomaly audience")
            if route.plant_id != alert.plant_id and not (
                route.audience == NotificationAudience.DATA_OPERATOR
                and route.plant_id == "*"
            ):
                raise ValueError("recipient plant_id does not match the anomaly plant_id")
            if route.recipient_id in seen:
                continue
            seen.add(route.recipient_id)
            results.append(
                self.outbox.enqueue(
                    NotificationEnvelope.for_anomaly(
                        alert,
                        route,
                        template_id=self.template_id,
                    ),
                    now=now,
                )
            )
        return results


class NotificationDispatcher:
    """Claim outbox work and dispatch it through runtime-only contacts/providers."""

    def __init__(
        self,
        settings: NotificationSettings,
        outbox: NotificationOutbox,
        contact_directory: ContactDirectory,
        providers: Mapping[NotificationChannel | str, NotificationProvider],
    ):
        self.settings = settings
        self.outbox = outbox
        self.contact_directory = contact_directory
        self.providers = {
            channel
            if isinstance(channel, NotificationChannel)
            else NotificationChannel(channel): provider
            for channel, provider in providers.items()
        }

    def run_once(self, *, now: datetime | None = None) -> DispatchSummary:
        records = self.outbox.claim_due(
            limit=self.settings.batch_size,
            lease_seconds=self.settings.lease_seconds,
            now=now,
        )
        counts = {
            "candidates": len(records),
            "accepted": 0,
            "suppressed": 0,
            "retry_required": 0,
            "fallback_scheduled": 0,
            "dead_lettered": 0,
            "in_doubt": 0,
        }
        for record in records:
            if not self.settings.external_delivery_allowed:
                self.outbox.mark_dry_run(
                    record.notification_id,
                    expected_attempt=record.attempts_total,
                    now=now,
                )
                counts["suppressed"] += 1
                continue
            try:
                provider = self.providers.get(record.channel)
                if provider is None:
                    raise PermanentNotificationError(
                        f"no provider is configured for channel={record.channel.value}"
                    )
                address = self.contact_directory.resolve(
                    record.recipient_id, record.channel
                )
                result = provider.send(
                    ProviderMessage(
                        event_id=record.event_id,
                        plant_id=record.plant_id,
                        recipient_id=record.recipient_id,
                        recipient_address=address,
                        channel=record.channel,
                        template_id=record.template_id,
                        subject=record.subject,
                        body=record.body,
                        metadata=record.metadata,
                        idempotency_key=record.idempotency_key,
                    )
                )
                # A SOLAPI HTTP 2xx means provider acceptance, not handset delivery.
                self.outbox.mark_accepted(
                    record.notification_id,
                    expected_attempt=record.attempts_total,
                    provider_message_id=result.provider_message_id,
                    now=now,
                )
                counts["accepted"] += 1
            except AmbiguousNotificationError as exc:
                self.outbox.mark_in_doubt(
                    record.notification_id,
                    expected_attempt=record.attempts_total,
                    reason=str(exc),
                    now=now,
                )
                counts["in_doubt"] += 1
            except (PermanentNotificationError, TransientNotificationError) as exc:
                transition = self.outbox.mark_failed(
                    record.notification_id,
                    expected_attempt=record.attempts_total,
                    error=str(exc),
                    transient=isinstance(exc, TransientNotificationError),
                    max_attempts_per_channel=self.settings.max_attempts_per_channel,
                    retry_base_seconds=self.settings.retry_base_seconds,
                    retry_max_seconds=self.settings.retry_max_seconds,
                    now=now,
                )
                if transition.status == NotificationStatus.RETRY:
                    counts["retry_required"] += 1
                if transition.fallback_used:
                    counts["fallback_scheduled"] += 1
                if transition.status == NotificationStatus.DEAD_LETTER:
                    counts["dead_lettered"] += 1
            except Exception as exc:
                # Unknown adapter outcomes are quarantined. Retrying here could
                # duplicate a message that the provider accepted before failing.
                self.outbox.mark_in_doubt(
                    record.notification_id,
                    expected_attempt=record.attempts_total,
                    reason=f"unexpected provider outcome: {type(exc).__name__}",
                    now=now,
                )
                counts["in_doubt"] += 1
        return DispatchSummary(**counts)


def build_solapi_dispatcher(
    settings: NotificationSettings,
    outbox: NotificationOutbox,
    contact_directory: ContactDirectory,
    solapi: SolapiSettings,
    *,
    session: object | None = None,
) -> NotificationDispatcher:
    """Construct the two-channel SOLAPI adapter after the caller passes --live."""

    providers = {
        channel: SolapiProvider(channel, solapi, session=session)
        for channel in (
            NotificationChannel.KAKAO_ALIMTALK,
            NotificationChannel.SMS,
        )
    }
    return NotificationDispatcher(settings, outbox, contact_directory, providers)
