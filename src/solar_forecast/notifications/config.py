from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import datetime, timezone
from hashlib import sha256
import hmac
import os
from pathlib import Path
import secrets
from typing import Mapping
from urllib.parse import urlparse


def _parse_bool(name: str, value: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off"}:
        return False
    raise ValueError(f"{name} must be a boolean value")


def _positive_int(name: str, value: str | int) -> int:
    result = int(value)
    if result < 1:
        raise ValueError(f"{name} must be positive")
    return result


@dataclass(frozen=True)
class NotificationSettings:
    """Operational controls; external delivery is fail-safe off by default."""

    enabled: bool = False
    dry_run: bool = True
    live_confirmation: bool = False
    outbox_path: Path = Path("artifacts/notifications/outbox.sqlite3")
    max_attempts_per_channel: int = 3
    retry_base_seconds: int = 60
    retry_max_seconds: int = 3600
    lease_seconds: int = 120
    batch_size: int = 50

    def __post_init__(self) -> None:
        for name in (
            "max_attempts_per_channel",
            "retry_base_seconds",
            "retry_max_seconds",
            "lease_seconds",
            "batch_size",
        ):
            object.__setattr__(self, name, _positive_int(name, getattr(self, name)))
        if self.retry_max_seconds < self.retry_base_seconds:
            raise ValueError("retry_max_seconds cannot be smaller than retry_base_seconds")

    @property
    def external_delivery_allowed(self) -> bool:
        return self.enabled and not self.dry_run and self.live_confirmation

    def confirm_live(self, *, confirmed: bool) -> "NotificationSettings":
        """Apply the separate CLI/API live gate; environment variables cannot set it."""

        return replace(self, live_confirmation=bool(confirmed))

    @classmethod
    def from_env(
        cls,
        environ: Mapping[str, str] | None = None,
        *,
        prefix: str = "SOLAR_NOTIFY_",
    ) -> "NotificationSettings":
        values = os.environ if environ is None else environ

        def item(name: str, default: str) -> str:
            return str(values.get(f"{prefix}{name}", default))

        return cls(
            enabled=_parse_bool(f"{prefix}ENABLED", item("ENABLED", "false")),
            dry_run=_parse_bool(f"{prefix}DRY_RUN", item("DRY_RUN", "true")),
            outbox_path=Path(
                item("OUTBOX_PATH", "artifacts/notifications/outbox.sqlite3")
            ),
            max_attempts_per_channel=_positive_int(
                f"{prefix}MAX_ATTEMPTS_PER_CHANNEL",
                item("MAX_ATTEMPTS_PER_CHANNEL", "3"),
            ),
            retry_base_seconds=_positive_int(
                f"{prefix}RETRY_BASE_SECONDS", item("RETRY_BASE_SECONDS", "60")
            ),
            retry_max_seconds=_positive_int(
                f"{prefix}RETRY_MAX_SECONDS", item("RETRY_MAX_SECONDS", "3600")
            ),
            lease_seconds=_positive_int(
                f"{prefix}LEASE_SECONDS", item("LEASE_SECONDS", "120")
            ),
            batch_size=_positive_int(f"{prefix}BATCH_SIZE", item("BATCH_SIZE", "50")),
        )

    def describe(self) -> dict[str, object]:
        return {
            "enabled": self.enabled,
            "dry_run": self.dry_run,
            "live_confirmation": self.live_confirmation,
            "external_delivery_allowed": self.external_delivery_allowed,
            "outbox_path": str(self.outbox_path),
            "max_attempts_per_channel": self.max_attempts_per_channel,
            "retry_base_seconds": self.retry_base_seconds,
            "retry_max_seconds": self.retry_max_seconds,
            "lease_seconds": self.lease_seconds,
            "batch_size": self.batch_size,
        }


@dataclass(frozen=True)
class SolapiSettings:
    """SOLAPI credentials are runtime-only and hidden from repr/output."""

    api_key: str = field(repr=False)
    api_secret: str = field(repr=False)
    sender_number: str = field(repr=False)
    kakao_pf_id: str = field(repr=False)
    kakao_template_id: str
    endpoint: str = "https://api.solapi.com/messages/v4/send-many/detail"
    timeout_seconds: int = 10

    def __post_init__(self) -> None:
        parsed = urlparse(self.endpoint)
        if parsed.scheme != "https" or not parsed.netloc:
            raise ValueError("SOLAPI endpoint must be an absolute HTTPS URL")
        if (
            parsed.hostname != "api.solapi.com"
            or parsed.path != "/messages/v4/send-many/detail"
            or parsed.query
            or parsed.fragment
        ):
            raise ValueError(
                "SOLAPI credentials may only be sent to the official "
                "api.solapi.com/messages/v4/send-many/detail endpoint"
            )
        for name in (
            "api_key",
            "api_secret",
            "sender_number",
            "kakao_pf_id",
            "kakao_template_id",
        ):
            if not str(getattr(self, name)).strip():
                raise ValueError(f"{name} cannot be blank")
        object.__setattr__(
            self,
            "timeout_seconds",
            _positive_int("timeout_seconds", self.timeout_seconds),
        )

    @classmethod
    def from_env(
        cls,
        environ: Mapping[str, str] | None = None,
        *,
        prefix: str = "SOLAR_NOTIFY_",
    ) -> "SolapiSettings":
        values = os.environ if environ is None else environ

        def required(name: str) -> str:
            key = f"{prefix}{name}"
            value = str(values.get(key, "")).strip()
            if not value:
                raise ValueError(f"Missing runtime notification setting: {key}")
            return value

        return cls(
            api_key=required("SOLAPI_API_KEY"),
            api_secret=required("SOLAPI_API_SECRET"),
            sender_number=required("SOLAPI_SENDER_NUMBER"),
            kakao_pf_id=required("SOLAPI_KAKAO_PF_ID"),
            kakao_template_id=required("SOLAPI_KAKAO_TEMPLATE_ID"),
            endpoint=str(
                values.get(
                    f"{prefix}SOLAPI_ENDPOINT",
                    "https://api.solapi.com/messages/v4/send-many/detail",
                )
            ),
            timeout_seconds=_positive_int(
                f"{prefix}SOLAPI_HTTP_TIMEOUT_SECONDS",
                str(values.get(f"{prefix}SOLAPI_HTTP_TIMEOUT_SECONDS", "10")),
            ),
        )

    def describe(self) -> dict[str, object]:
        return {
            "provider": "solapi",
            "endpoint_host": urlparse(self.endpoint).netloc,
            "api_key_configured": bool(self.api_key),
            "api_secret_configured": bool(self.api_secret),
            "kakao_template_id": self.kakao_template_id,
            "sender_number_configured": bool(self.sender_number),
            "kakao_pf_id_configured": bool(self.kakao_pf_id),
            "timeout_seconds": self.timeout_seconds,
        }

    def authorization_header(
        self,
        *,
        now: datetime | None = None,
        salt: str | None = None,
    ) -> str:
        current = (now or datetime.now(timezone.utc)).astimezone(timezone.utc)
        date = current.isoformat(timespec="milliseconds").replace("+00:00", "Z")
        nonce = salt or secrets.token_hex(16)
        signature = hmac.new(
            self.api_secret.encode("utf-8"),
            f"{date}{nonce}".encode("utf-8"),
            sha256,
        ).hexdigest()
        return (
            f"HMAC-SHA256 apiKey={self.api_key}, date={date}, "
            f"salt={nonce}, signature={signature}"
        )
