from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any, Mapping, Protocol

from solar_forecast.anomalies import INTERPRETATION_LIMIT

from .config import SolapiSettings
from .models import DeliveryResult, NotificationChannel


class TransientNotificationError(RuntimeError):
    """A temporary provider failure that is safe to retry."""


class PermanentNotificationError(RuntimeError):
    """A request failure that should fall back or enter the dead-letter queue."""


class AmbiguousNotificationError(RuntimeError):
    """The provider may have accepted the request, so automatic retry is unsafe."""


@dataclass(frozen=True)
class ProviderMessage:
    event_id: str
    plant_id: str
    recipient_id: str
    recipient_address: str
    channel: NotificationChannel
    template_id: str
    subject: str
    body: str
    metadata: Mapping[str, Any]
    idempotency_key: str


class NotificationProvider(Protocol):
    channel: NotificationChannel

    def send(self, message: ProviderMessage) -> DeliveryResult:
        ...


class ContactDirectory(Protocol):
    def resolve(self, recipient_id: str, channel: NotificationChannel) -> str:
        """Resolve a manager reference at dispatch time without persisting phone numbers."""


class MappingContactDirectory:
    """In-memory adapter for a runtime-injected secret-manager/directory snapshot."""

    def __init__(self, contacts: Mapping[tuple[str, NotificationChannel | str], str]):
        self._contacts = {
            (
                str(recipient_id).strip(),
                channel if isinstance(channel, NotificationChannel) else NotificationChannel(channel),
            ): str(address).strip()
            for (recipient_id, channel), address in contacts.items()
        }

    def resolve(self, recipient_id: str, channel: NotificationChannel) -> str:
        address = self._contacts.get((recipient_id, channel), "")
        if not address:
            raise PermanentNotificationError(
                f"No runtime contact is registered for recipient={recipient_id}, channel={channel.value}"
            )
        return address


class SolapiProvider:
    """SOLAPI v4 adapter for approved Kakao AlimTalk templates and SMS/LMS."""

    def __init__(
        self,
        channel: NotificationChannel,
        settings: SolapiSettings,
        *,
        session: Any | None = None,
    ):
        self.channel = NotificationChannel(channel)
        self.settings = settings
        if session is None:
            import requests

            session = requests.Session()
        self._session = session

    def send(self, message: ProviderMessage) -> DeliveryResult:
        if message.channel != self.channel:
            raise PermanentNotificationError(
                f"Provider channel mismatch: {message.channel.value}"
            )
        outbound = {
            "to": _recipient_phone(message.recipient_address),
            "from": _sender_phone(self.settings.sender_number),
        }
        if message.channel == NotificationChannel.KAKAO_ALIMTALK:
            outbound["type"] = "ATA"
            outbound["kakaoOptions"] = {
                "pfId": self.settings.kakao_pf_id,
                "templateId": self.settings.kakao_template_id,
                "variables": self._template_variables(message),
                # Let SOLAPI perform its registered-number SMS replacement if
                # Kakao delivery fails after the API request was accepted.
                "disableSms": False,
            }
        else:
            outbound["type"] = (
                "SMS" if len(message.body.encode("utf-8")) <= 90 else "LMS"
            )
            outbound["text"] = message.body
            if outbound["type"] == "LMS":
                outbound["subject"] = message.subject[:40]
        payload = {"messages": [outbound]}
        headers = {
            "Authorization": self.settings.authorization_header(),
            "Content-Type": "application/json",
        }
        try:
            response = self._session.post(
                self.settings.endpoint,
                json=payload,
                headers=headers,
                timeout=self.settings.timeout_seconds,
            )
        except Exception as exc:
            raise AmbiguousNotificationError(
                f"SOLAPI request outcome is unknown: {type(exc).__name__}"
            ) from exc
        status = int(response.status_code)
        if status == 429:
            raise TransientNotificationError(f"SOLAPI returned HTTP {status}")
        if status == 408 or status >= 500:
            raise AmbiguousNotificationError(
                f"SOLAPI request outcome is unknown after HTTP {status}"
            )
        if status < 200 or status >= 300:
            raise PermanentNotificationError(f"SOLAPI returned HTTP {status}")
        try:
            result = response.json()
        except Exception as exc:
            raise AmbiguousNotificationError(
                "SOLAPI accepted the HTTP request but returned invalid JSON"
            ) from exc
        failures = result.get("failedMessageList") or []
        if failures:
            raise PermanentNotificationError("SOLAPI rejected the notification message")
        group = result.get("groupInfo") if isinstance(result.get("groupInfo"), dict) else {}
        message_id = str(
            result.get("groupId")
            or group.get("groupId")
            or result.get("messageId")
            or ""
        ).strip()
        if not message_id:
            raise AmbiguousNotificationError(
                "SOLAPI response did not include a group/message identifier"
            )
        return DeliveryResult(provider_message_id=message_id)

    @staticmethod
    def _template_variables(message: ProviderMessage) -> dict[str, str]:
        metadata = message.metadata
        severity = {
            "review": "검토",
            "high": "주의",
            "critical": "긴급",
        }.get(str(metadata.get("severity")), "검토")
        return {
            "#{지역}": str(metadata.get("region") or "지역 미확인"),
            "#{발전소명}": str(metadata.get("plant_name") or message.plant_id),
            "#{관측시각}": str(
                metadata.get("observed_at_kst") or metadata.get("observed_at") or ""
            ),
            "#{심각도}": severity,
            "#{판단근거}": str(
                metadata.get("influence_factor_label")
                or metadata.get("influence_factor")
                or ""
            ),
            "#{요약}": str(metadata.get("summary") or ""),
            "#{주의사항}": INTERPRETATION_LIMIT,
        }


def _recipient_phone(value: str) -> str:
    normalized = re.sub(r"[^0-9]", "", str(value))
    if len(normalized) not in {10, 11} or not normalized.startswith("01"):
        raise PermanentNotificationError(
            "recipient address is not a valid Korean mobile number"
        )
    return normalized


def _sender_phone(value: str) -> str:
    normalized = re.sub(r"[^0-9]", "", str(value))
    if len(normalized) not in {8, 9, 10, 11}:
        raise PermanentNotificationError(
            "SOLAPI sender number is not a valid registered Korean sender number"
        )
    return normalized
