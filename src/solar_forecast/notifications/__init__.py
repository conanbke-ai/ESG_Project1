"""Fail-safe anomaly notification outbox and SOLAPI delivery adapters."""

from .config import NotificationSettings, SolapiSettings
from .models import (
    NOTIFICATION_CONTRACT,
    OPERATIONAL_EVENT_CONTRACT,
    AnomalyAlert,
    DeliveryResult,
    NotificationAudience,
    NotificationChannel,
    NotificationStatus,
    RecipientRoute,
    Severity,
    stable_event_id,
)
from .outbox import (
    OUTBOX_SCHEMA_VERSION,
    EnqueueResult,
    FailureTransition,
    NotificationOutbox,
    OutboxRecord,
)
from .providers import (
    AmbiguousNotificationError,
    ContactDirectory,
    MappingContactDirectory,
    NotificationProvider,
    PermanentNotificationError,
    ProviderMessage,
    SolapiProvider,
    TransientNotificationError,
)
from .runtime import ROUTE_DIRECTORY_CONTRACT, RuntimeRouteDirectory
from .service import (
    DispatchSummary,
    NotificationDispatcher,
    NotificationService,
    build_solapi_dispatcher,
)

__all__ = [
    "NOTIFICATION_CONTRACT",
    "OPERATIONAL_EVENT_CONTRACT",
    "OUTBOX_SCHEMA_VERSION",
    "ROUTE_DIRECTORY_CONTRACT",
    "AnomalyAlert",
    "AmbiguousNotificationError",
    "ContactDirectory",
    "DeliveryResult",
    "DispatchSummary",
    "EnqueueResult",
    "FailureTransition",
    "MappingContactDirectory",
    "NotificationAudience",
    "NotificationChannel",
    "NotificationDispatcher",
    "NotificationOutbox",
    "NotificationProvider",
    "NotificationService",
    "NotificationSettings",
    "NotificationStatus",
    "OutboxRecord",
    "PermanentNotificationError",
    "ProviderMessage",
    "RecipientRoute",
    "RuntimeRouteDirectory",
    "Severity",
    "SolapiProvider",
    "SolapiSettings",
    "TransientNotificationError",
    "build_solapi_dispatcher",
    "stable_event_id",
]
