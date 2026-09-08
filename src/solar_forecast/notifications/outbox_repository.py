"""SQLite outbox의 멱등성·lease·재시도·완료 상태 영속화."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from hashlib import sha256
import json
from pathlib import Path
import sqlite3
from typing import Any, Mapping, Sequence

from solar_forecast.notifications.contracts import (
    NOTIFICATION_CONTRACT,
    NotificationChannel,
    NotificationEnvelope,
    NotificationStatus,
    ensure_utc,
    redact_sensitive,
    utc_now,
)


OUTBOX_SCHEMA_VERSION = "solar-notification-outbox.v1"


@dataclass(frozen=True)
class EnqueueResult:
    notification_id: int
    created: bool


@dataclass(frozen=True)
class OutboxRecord:
    notification_id: int
    event_id: str
    plant_id: str
    recipient_id: str
    channels: tuple[NotificationChannel, ...]
    channel_index: int
    template_id: str
    subject: str
    body: str
    metadata: Mapping[str, Any]
    status: NotificationStatus
    attempts_total: int
    attempts_on_channel: int
    available_at: datetime
    lease_until: datetime | None
    provider_message_id: str | None
    last_error: str | None
    idempotency_key: str

    @property
    def channel(self) -> NotificationChannel:
        return self.channels[self.channel_index]


@dataclass(frozen=True)
class FailureTransition:
    notification_id: int
    status: NotificationStatus
    channel: NotificationChannel
    fallback_used: bool
    available_at: datetime | None


class NotificationOutbox:
    """Transactional SQLite outbox with leases, fallback and append-only audit."""

    def __init__(self, path: Path):
        self.path = Path(path)
        if str(self.path) == ":memory:":
            raise ValueError("use a file-backed SQLite outbox so state survives restarts")
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    def enqueue(
        self,
        envelope: NotificationEnvelope,
        *,
        now: datetime | None = None,
    ) -> EnqueueResult:
        timestamp = _format_utc_timestamp(now or utc_now())
        channel_values = [channel.value for channel in envelope.channels]
        routing_key = self._routing_key(envelope.event_id, envelope.recipient_id)
        with self._connection() as connection:
            connection.execute("BEGIN IMMEDIATE")
            cursor = connection.execute(
                """
                INSERT INTO notifications (
                    routing_key, event_id, plant_id, recipient_id, channels_json,
                    current_channel_index, template_id, subject, body, metadata_json,
                    status, attempts_total, attempts_on_channel, available_at,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, 0, ?, ?, ?, ?, ?, 0, 0, ?, ?, ?)
                ON CONFLICT(routing_key) DO NOTHING
                """,
                (
                    routing_key,
                    envelope.event_id,
                    envelope.plant_id,
                    envelope.recipient_id,
                    _serialize_outbox_payload(channel_values),
                    envelope.template_id,
                    envelope.subject,
                    envelope.body,
                    _serialize_outbox_payload(dict(envelope.metadata)),
                    NotificationStatus.PENDING.value,
                    timestamp,
                    timestamp,
                    timestamp,
                ),
            )
            created = cursor.rowcount == 1
            row = connection.execute(
                "SELECT id FROM notifications WHERE routing_key = ?", (routing_key,)
            ).fetchone()
            notification_id = int(row["id"])
            if created:
                for channel in envelope.channels:
                    connection.execute(
                        """
                        INSERT INTO delivery_keys (
                            notification_id, event_id, recipient_id, channel,
                            idempotency_key
                        ) VALUES (?, ?, ?, ?, ?)
                        """,
                        (
                            notification_id,
                            envelope.event_id,
                            envelope.recipient_id,
                            channel.value,
                            self._delivery_key(
                                envelope.event_id, envelope.recipient_id, channel
                            ),
                        ),
                    )
                self._audit(
                    connection,
                    notification_id,
                    occurred_at=timestamp,
                    action="enqueued",
                    from_status=None,
                    to_status=NotificationStatus.PENDING,
                    attempt=0,
                    detail={"channels": channel_values},
                )
            connection.commit()
        return EnqueueResult(notification_id=notification_id, created=created)

    def claim_due(
        self,
        *,
        limit: int,
        lease_seconds: int,
        now: datetime | None = None,
    ) -> list[OutboxRecord]:
        if limit < 1 or lease_seconds < 1:
            raise ValueError("claim limit and lease_seconds must be positive")
        current = ensure_utc(now or utc_now())
        current_iso = _format_utc_timestamp(current)
        lease_until = _format_utc_timestamp(current + timedelta(seconds=lease_seconds))
        claimed: list[int] = []
        with self._connection() as connection:
            connection.execute("BEGIN IMMEDIATE")
            expired = connection.execute(
                """
                SELECT id, attempts_total FROM notifications
                 WHERE status = ? AND lease_until <= ?
                """,
                (NotificationStatus.IN_PROGRESS.value, current_iso),
            ).fetchall()
            for row in expired:
                connection.execute(
                    """
                    UPDATE notifications
                       SET status = ?, lease_until = NULL, updated_at = ?,
                           last_error = ?
                     WHERE id = ? AND status = ?
                    """,
                    (
                        NotificationStatus.IN_DOUBT.value,
                        current_iso,
                        "worker lease expired after dispatch claim; manual provider reconciliation required",
                        int(row["id"]),
                        NotificationStatus.IN_PROGRESS.value,
                    ),
                )
                self._audit(
                    connection,
                    int(row["id"]),
                    occurred_at=current_iso,
                    action="lease_expired_in_doubt",
                    from_status=NotificationStatus.IN_PROGRESS,
                    to_status=NotificationStatus.IN_DOUBT,
                    attempt=int(row["attempts_total"]),
                    detail={"automatic_resend": False},
                )
            rows = connection.execute(
                """
                SELECT id FROM notifications
                 WHERE status IN (?, ?) AND available_at <= ?
                 ORDER BY available_at, id
                 LIMIT ?
                """,
                (
                    NotificationStatus.PENDING.value,
                    NotificationStatus.RETRY.value,
                    current_iso,
                    int(limit),
                ),
            ).fetchall()
            for row in rows:
                notification_id = int(row["id"])
                before = connection.execute(
                    "SELECT status, attempts_total FROM notifications WHERE id = ?",
                    (notification_id,),
                ).fetchone()
                connection.execute(
                    """
                    UPDATE notifications
                       SET status = ?, attempts_total = attempts_total + 1,
                           attempts_on_channel = attempts_on_channel + 1,
                           lease_until = ?, updated_at = ?
                     WHERE id = ?
                    """,
                    (
                        NotificationStatus.IN_PROGRESS.value,
                        lease_until,
                        current_iso,
                        notification_id,
                    ),
                )
                claimed.append(notification_id)
                self._audit(
                    connection,
                    notification_id,
                    occurred_at=current_iso,
                    action="claimed",
                    from_status=NotificationStatus(before["status"]),
                    to_status=NotificationStatus.IN_PROGRESS,
                    attempt=int(before["attempts_total"]) + 1,
                    detail={"lease_until": lease_until},
                )
            connection.commit()
        return [self.get(notification_id) for notification_id in claimed]

    def mark_accepted(
        self,
        notification_id: int,
        *,
        expected_attempt: int,
        provider_message_id: str,
        now: datetime | None = None,
    ) -> None:
        self._finish(
            notification_id,
            expected_attempt=expected_attempt,
            target=NotificationStatus.ACCEPTED,
            action="provider_accepted",
            provider_message_id=provider_message_id,
            now=now,
        )

    def mark_dry_run(
        self,
        notification_id: int,
        *,
        expected_attempt: int,
        now: datetime | None = None,
    ) -> None:
        self._finish(
            notification_id,
            expected_attempt=expected_attempt,
            target=NotificationStatus.DRY_RUN,
            action="suppressed_dry_run",
            provider_message_id=None,
            now=now,
        )

    def mark_in_doubt(
        self,
        notification_id: int,
        *,
        expected_attempt: int,
        reason: str,
        now: datetime | None = None,
    ) -> None:
        """Quarantine an ambiguous provider outcome without automatic resend."""

        timestamp = _format_utc_timestamp(now or utc_now())
        safe_reason = redact_sensitive(reason).strip() or "provider outcome is unknown"
        with self._connection() as connection:
            connection.execute("BEGIN IMMEDIATE")
            self._locked_in_progress(connection, notification_id, expected_attempt)
            connection.execute(
                """
                UPDATE notifications
                   SET status = ?, lease_until = NULL, last_error = ?, updated_at = ?
                 WHERE id = ?
                """,
                (
                    NotificationStatus.IN_DOUBT.value,
                    safe_reason,
                    timestamp,
                    notification_id,
                ),
            )
            self._audit(
                connection,
                notification_id,
                occurred_at=timestamp,
                action="provider_outcome_unknown",
                from_status=NotificationStatus.IN_PROGRESS,
                to_status=NotificationStatus.IN_DOUBT,
                attempt=expected_attempt,
                detail={
                    "automatic_resend": False,
                    "reason": safe_reason,
                },
            )
            connection.commit()

    def mark_failed(
        self,
        notification_id: int,
        *,
        expected_attempt: int,
        error: str,
        transient: bool,
        max_attempts_per_channel: int,
        retry_base_seconds: int,
        retry_max_seconds: int,
        now: datetime | None = None,
    ) -> FailureTransition:
        if min(max_attempts_per_channel, retry_base_seconds, retry_max_seconds) < 1:
            raise ValueError("retry policy values must be positive")
        current = ensure_utc(now or utc_now())
        current_iso = _format_utc_timestamp(current)
        safe_error = redact_sensitive(error)
        with self._connection() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = self._locked_in_progress(connection, notification_id, expected_attempt)
            channels = tuple(NotificationChannel(item) for item in json.loads(row["channels_json"]))
            index = int(row["current_channel_index"])
            channel_attempts = int(row["attempts_on_channel"])
            fallback = False
            if transient and channel_attempts < max_attempts_per_channel:
                target = NotificationStatus.RETRY
                delay = min(
                    retry_max_seconds,
                    retry_base_seconds * (2 ** max(0, channel_attempts - 1)),
                )
                available = current + timedelta(seconds=delay)
                next_index = index
                next_channel_attempts = channel_attempts
                action = "retry_scheduled"
            elif index + 1 < len(channels):
                target = NotificationStatus.RETRY
                available = current
                next_index = index + 1
                next_channel_attempts = 0
                fallback = True
                action = "fallback_scheduled"
            else:
                target = NotificationStatus.DEAD_LETTER
                available = current
                next_index = index
                next_channel_attempts = channel_attempts
                action = "dead_lettered"
            connection.execute(
                """
                UPDATE notifications
                   SET status = ?, current_channel_index = ?,
                       attempts_on_channel = ?, available_at = ?, lease_until = NULL,
                       last_error = ?, updated_at = ?
                 WHERE id = ?
                """,
                (
                    target.value,
                    next_index,
                    next_channel_attempts,
                    _format_utc_timestamp(available),
                    safe_error,
                    current_iso,
                    notification_id,
                ),
            )
            self._audit(
                connection,
                notification_id,
                occurred_at=current_iso,
                action=action,
                from_status=NotificationStatus.IN_PROGRESS,
                to_status=target,
                attempt=expected_attempt,
                detail={
                    "failed_channel": channels[index].value,
                    "next_channel": channels[next_index].value if fallback else None,
                    "transient": bool(transient),
                    "error": safe_error,
                },
            )
            connection.commit()
        return FailureTransition(
            notification_id=notification_id,
            status=target,
            channel=channels[next_index],
            fallback_used=fallback,
            available_at=available if target == NotificationStatus.RETRY else None,
        )

    def resolve_in_doubt(
        self,
        notification_id: int,
        *,
        resolution_note: str,
        provider_message_id: str | None = None,
        retry: bool = False,
        now: datetime | None = None,
    ) -> None:
        if bool(provider_message_id) == bool(retry):
            raise ValueError("choose exactly one resolution: provider_message_id or retry")
        note = redact_sensitive(resolution_note).strip()
        if not note:
            raise ValueError("resolution_note is required")
        timestamp = _format_utc_timestamp(now or utc_now())
        target = NotificationStatus.RETRY if retry else NotificationStatus.ACCEPTED
        with self._connection() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                "SELECT status, attempts_total FROM notifications WHERE id = ?",
                (notification_id,),
            ).fetchone()
            if row is None or row["status"] != NotificationStatus.IN_DOUBT.value:
                raise ValueError("notification is not in_doubt")
            connection.execute(
                """
                UPDATE notifications
                   SET status = ?, provider_message_id = ?, available_at = ?,
                       lease_until = NULL, last_error = NULL, updated_at = ?
                 WHERE id = ?
                """,
                (target.value, provider_message_id, timestamp, timestamp, notification_id),
            )
            self._audit(
                connection,
                notification_id,
                occurred_at=timestamp,
                action="in_doubt_resolved",
                from_status=NotificationStatus.IN_DOUBT,
                to_status=target,
                attempt=int(row["attempts_total"]),
                detail={"resolution_note": note, "manual_retry": bool(retry)},
            )
            connection.commit()

    def get(self, notification_id: int) -> OutboxRecord:
        with self._connection() as connection:
            row = connection.execute(
                """
                SELECT n.*, d.idempotency_key
                  FROM notifications AS n
                  JOIN delivery_keys AS d
                    ON d.notification_id = n.id
                   AND d.channel = json_extract(n.channels_json, '$[' || n.current_channel_index || ']')
                 WHERE n.id = ?
                """,
                (notification_id,),
            ).fetchone()
        if row is None:
            raise KeyError(f"unknown notification id: {notification_id}")
        return self._record(row)

    def audit_events(self, notification_id: int) -> list[dict[str, Any]]:
        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT occurred_at, action, from_status, to_status, attempt, detail_json
                  FROM notification_audit
                 WHERE notification_id = ? ORDER BY id
                """,
                (notification_id,),
            ).fetchall()
        return [
            {
                "occurred_at": row["occurred_at"],
                "action": row["action"],
                "from_status": row["from_status"],
                "to_status": row["to_status"],
                "attempt": int(row["attempt"]),
                "detail": json.loads(row["detail_json"]),
            }
            for row in rows
        ]

    def status_counts(self) -> dict[str, int]:
        with self._connection() as connection:
            rows = connection.execute(
                "SELECT status, COUNT(*) AS count FROM notifications GROUP BY status"
            ).fetchall()
        return {str(row["status"]): int(row["count"]) for row in rows}

    def _finish(
        self,
        notification_id: int,
        *,
        expected_attempt: int,
        target: NotificationStatus,
        action: str,
        provider_message_id: str | None,
        now: datetime | None,
    ) -> None:
        timestamp = _format_utc_timestamp(now or utc_now())
        with self._connection() as connection:
            connection.execute("BEGIN IMMEDIATE")
            self._locked_in_progress(connection, notification_id, expected_attempt)
            connection.execute(
                """
                UPDATE notifications
                   SET status = ?, lease_until = NULL, provider_message_id = ?,
                       last_error = NULL, updated_at = ?
                 WHERE id = ?
                """,
                (target.value, provider_message_id, timestamp, notification_id),
            )
            self._audit(
                connection,
                notification_id,
                occurred_at=timestamp,
                action=action,
                from_status=NotificationStatus.IN_PROGRESS,
                to_status=target,
                attempt=expected_attempt,
                detail={"provider_reference_recorded": bool(provider_message_id)},
            )
            connection.commit()

    @staticmethod
    def _locked_in_progress(
        connection: sqlite3.Connection,
        notification_id: int,
        expected_attempt: int,
    ) -> sqlite3.Row:
        row = connection.execute(
            "SELECT * FROM notifications WHERE id = ?", (notification_id,)
        ).fetchone()
        if row is None:
            raise KeyError(f"unknown notification id: {notification_id}")
        if (
            row["status"] != NotificationStatus.IN_PROGRESS.value
            or int(row["attempts_total"]) != expected_attempt
        ):
            raise ValueError("notification claim is stale or no longer in progress")
        return row

    def _initialize(self) -> None:
        with self._connection() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS outbox_metadata (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS notifications (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    routing_key TEXT NOT NULL UNIQUE,
                    event_id TEXT NOT NULL,
                    plant_id TEXT NOT NULL,
                    recipient_id TEXT NOT NULL,
                    channels_json TEXT NOT NULL,
                    current_channel_index INTEGER NOT NULL,
                    template_id TEXT NOT NULL,
                    subject TEXT NOT NULL,
                    body TEXT NOT NULL,
                    metadata_json TEXT NOT NULL,
                    status TEXT NOT NULL,
                    attempts_total INTEGER NOT NULL,
                    attempts_on_channel INTEGER NOT NULL,
                    available_at TEXT NOT NULL,
                    lease_until TEXT,
                    provider_message_id TEXT,
                    last_error TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_notifications_due
                    ON notifications(status, available_at);
                CREATE TABLE IF NOT EXISTS delivery_keys (
                    notification_id INTEGER NOT NULL REFERENCES notifications(id),
                    event_id TEXT NOT NULL,
                    recipient_id TEXT NOT NULL,
                    channel TEXT NOT NULL,
                    idempotency_key TEXT NOT NULL UNIQUE,
                    PRIMARY KEY(notification_id, channel),
                    UNIQUE(event_id, recipient_id, channel)
                );
                CREATE TABLE IF NOT EXISTS notification_audit (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    notification_id INTEGER NOT NULL REFERENCES notifications(id),
                    occurred_at TEXT NOT NULL,
                    action TEXT NOT NULL,
                    from_status TEXT,
                    to_status TEXT NOT NULL,
                    attempt INTEGER NOT NULL,
                    detail_json TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_notification_audit_message
                    ON notification_audit(notification_id, id);
                """
            )
            existing = connection.execute(
                "SELECT value FROM outbox_metadata WHERE key = 'schema_version'"
            ).fetchone()
            if existing is not None and existing["value"] != OUTBOX_SCHEMA_VERSION:
                raise RuntimeError("notification outbox schema is incompatible")
            connection.execute(
                "INSERT OR IGNORE INTO outbox_metadata(key, value) VALUES('schema_version', ?)",
                (OUTBOX_SCHEMA_VERSION,),
            )

    def _connection(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path, timeout=30)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA busy_timeout = 30000")
        connection.execute("PRAGMA journal_mode = WAL")
        return connection

    @staticmethod
    def _record(row: sqlite3.Row) -> OutboxRecord:
        channels = tuple(NotificationChannel(value) for value in json.loads(row["channels_json"]))
        return OutboxRecord(
            notification_id=int(row["id"]),
            event_id=str(row["event_id"]),
            plant_id=str(row["plant_id"]),
            recipient_id=str(row["recipient_id"]),
            channels=channels,
            channel_index=int(row["current_channel_index"]),
            template_id=str(row["template_id"]),
            subject=str(row["subject"]),
            body=str(row["body"]),
            metadata=json.loads(row["metadata_json"]),
            status=NotificationStatus(row["status"]),
            attempts_total=int(row["attempts_total"]),
            attempts_on_channel=int(row["attempts_on_channel"]),
            available_at=_parse_utc_timestamp(row["available_at"]),
            lease_until=_parse_utc_timestamp(row["lease_until"]) if row["lease_until"] else None,
            provider_message_id=row["provider_message_id"],
            last_error=row["last_error"],
            idempotency_key=str(row["idempotency_key"]),
        )

    @staticmethod
    def _audit(
        connection: sqlite3.Connection,
        notification_id: int,
        *,
        occurred_at: str,
        action: str,
        from_status: NotificationStatus | None,
        to_status: NotificationStatus,
        attempt: int,
        detail: Mapping[str, Any],
    ) -> None:
        connection.execute(
            """
            INSERT INTO notification_audit(
                notification_id, occurred_at, action, from_status,
                to_status, attempt, detail_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                notification_id,
                occurred_at,
                action,
                from_status.value if from_status else None,
                to_status.value,
                int(attempt),
                _serialize_outbox_payload(dict(detail)),
            ),
        )

    @staticmethod
    def _routing_key(event_id: str, recipient_id: str) -> str:
        return sha256(
            f"{NOTIFICATION_CONTRACT}\x1f{event_id}\x1f{recipient_id}".encode("utf-8")
        ).hexdigest()

    @staticmethod
    def _delivery_key(
        event_id: str,
        recipient_id: str,
        channel: NotificationChannel,
    ) -> str:
        return sha256(
            (
                f"{NOTIFICATION_CONTRACT}\x1f{event_id}\x1f"
                f"{recipient_id}\x1f{channel.value}"
            ).encode("utf-8")
        ).hexdigest()


def _format_utc_timestamp(value: datetime) -> str:
    return ensure_utc(value).isoformat(timespec="microseconds")


def _parse_utc_timestamp(value: str) -> datetime:
    return datetime.fromisoformat(value).astimezone(timezone.utc)


def _serialize_outbox_payload(value: Mapping[str, Any] | Sequence[Any]) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        default=str,
    )
