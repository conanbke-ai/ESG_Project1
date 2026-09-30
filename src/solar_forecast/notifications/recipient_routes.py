"""로컬 수신 경로에서 연락처 환경변수 이름과 대상 매핑 해석."""
from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path
import re
from typing import Mapping

from solar_forecast.notifications.contracts import (
    NotificationAudience,
    NotificationChannel,
    RecipientRoute,
    normalize_channels,
)
from solar_forecast.notifications.delivery_provider import MappingContactDirectory


ROUTE_DIRECTORY_CONTRACT = "solar-notification-routes.v1"
_ENV_NAME = re.compile(r"^[A-Z][A-Z0-9_]{2,}$")


@dataclass(frozen=True)
class RuntimeRouteDirectory:
    """Routes public event keys to runtime-only manager phone environment values."""

    _routes: Mapping[str, tuple[RecipientRoute, ...]]
    contacts: MappingContactDirectory

    @classmethod
    def from_file(
        cls,
        path: Path,
        *,
        environ: Mapping[str, str] | None = None,
        require_contacts: bool,
    ) -> "RuntimeRouteDirectory":
        path = Path(path)
        if not path.is_file():
            raise FileNotFoundError(f"Notification route directory is missing: {path}")
        payload = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            raise ValueError("notification route directory must be a JSON object")
        if payload.get("schema_version") != ROUTE_DIRECTORY_CONTRACT:
            raise ValueError(f"schema_version must be {ROUTE_DIRECTORY_CONTRACT}")
        rows = payload.get("routes")
        if not isinstance(rows, list):
            raise ValueError("notification route directory requires a routes array")
        environment = os.environ if environ is None else environ
        grouped: dict[str, list[RecipientRoute]] = {}
        contacts: dict[tuple[str, NotificationChannel], str] = {}
        seen: set[tuple[str, str]] = set()
        for index, raw in enumerate(rows, start=1):
            if not isinstance(raw, dict):
                raise ValueError(f"notification route #{index} must be an object")
            if raw.get("active", True) is not True:
                continue
            forbidden = {"phone", "phone_number", "mobile", "contact"}.intersection(raw)
            if forbidden:
                raise ValueError(
                    "route files cannot contain phone numbers; use phone_env instead"
                )
            route_key = _required_text("route_key", raw.get("route_key"))
            recipient_id = _required_text("recipient_id", raw.get("recipient_id"))
            pair = (route_key, recipient_id)
            if pair in seen:
                raise ValueError(
                    f"duplicate notification route for route_key={route_key}, "
                    f"recipient_id={recipient_id}"
                )
            seen.add(pair)
            phone_env = _required_text("phone_env", raw.get("phone_env"))
            if not _ENV_NAME.fullmatch(phone_env):
                raise ValueError(f"invalid phone_env name for route #{index}: {phone_env}")
            channels = normalize_channels(
                raw.get("channels")
                or [
                    NotificationChannel.KAKAO_ALIMTALK.value,
                    NotificationChannel.SMS.value,
                ]
            )
            route = RecipientRoute(
                route_key=route_key,
                plant_id=_required_text("plant_id", raw.get("plant_id")),
                recipient_id=recipient_id,
                audience=NotificationAudience(
                    _required_text("audience", raw.get("audience"))
                ),
                channels=channels,
            )
            grouped.setdefault(route_key, []).append(route)
            phone = str(environment.get(phone_env, "")).strip()
            if require_contacts and not phone:
                raise ValueError(
                    f"Missing runtime manager contact environment variable: {phone_env}"
                )
            if phone:
                for channel in channels:
                    contacts[(recipient_id, channel)] = phone
        if not grouped:
            raise ValueError("notification route directory has no active routes")
        return cls(
            _routes={key: tuple(values) for key, values in grouped.items()},
            contacts=MappingContactDirectory(contacts),
        )

    def routes_for(
        self,
        route_key: str,
        audience: NotificationAudience,
        plant_id: str,
    ) -> tuple[RecipientRoute, ...]:
        plant_id = _required_text("plant_id", plant_id)
        routes = tuple(
            route
            for route in self._routes.get(str(route_key), ())
            if route.audience == audience
            and (
                route.plant_id == plant_id
                or (
                    route.audience == NotificationAudience.DATA_OPERATOR
                    and route.plant_id == "*"
                )
            )
        )
        if not routes:
            raise ValueError(
                f"No active notification route for route_key={route_key}, "
                f"audience={audience.value}, plant_id={plant_id}"
            )
        return routes


def _required_text(name: str, value: object) -> str:
    text = str(value).strip() if value is not None else ""
    if not text:
        raise ValueError(f"{name} cannot be blank")
    return text


__all__ = ["ROUTE_DIRECTORY_CONTRACT", "RuntimeRouteDirectory"]
