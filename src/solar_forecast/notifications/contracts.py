"""알림 이벤트·채널·상태·수신자 데이터 구조와 식별자 계약."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from hashlib import sha256
import json
import math
import re
from typing import Any, Mapping, Sequence
from zoneinfo import ZoneInfo

from solar_forecast.anomalies.event_batch import EVENT_CONTRACT, stable_operational_event_id
from solar_forecast.anomalies.influence_policy import (
    INTERPRETATION_LIMIT,
    validate_influence_factor,
)


NOTIFICATION_CONTRACT = "solar-anomaly-notification.v1"
OPERATIONAL_EVENT_CONTRACT = EVENT_CONTRACT


class NotificationChannel(str, Enum):
    KAKAO_ALIMTALK = "kakao_alimtalk"
    SMS = "sms"


class NotificationStatus(str, Enum):
    PENDING = "pending"
    IN_PROGRESS = "in_progress"
    RETRY = "retry"
    ACCEPTED = "accepted"
    DRY_RUN = "dry_run"
    DEAD_LETTER = "dead_letter"
    IN_DOUBT = "in_doubt"


class Severity(str, Enum):
    REVIEW = "review"
    HIGH = "high"
    CRITICAL = "critical"


class NotificationAudience(str, Enum):
    PLANT_MANAGER = "plant_manager"
    DATA_OPERATOR = "data_operator"


_FACTOR_LABELS = {
    "weather_impact": "기상 영향 가능성",
    "lower_than_expected": "예측 대비 발전량 저하",
    "higher_than_expected": "예측 대비 발전량 증가",
    "rapid_output_change": "급격한 출력 변화",
    "generation_stop_suspected": "발전 정지 의심 신호",
    "data_quality_issue": "데이터 품질 이상",
    "plant_specific_unknown_factor": "발전소별 미확인 요인",
    "unknown_external_factor": "미확인 외부 요인",
}


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def ensure_utc(value: datetime) -> datetime:
    if value.tzinfo is None or value.utcoffset() is None:
        raise ValueError("datetime values must include a timezone")
    return value.astimezone(timezone.utc)


def stable_event_id(
    *,
    model: str,
    plant_id: str,
    observed_at: datetime,
    detector_version: str,
    influence_factor: str,
    signal_type: str,
    detector_role: str,
) -> str:
    """Build a deterministic detector event ID suitable for idempotent reruns."""

    return stable_operational_event_id(
        detector_id=model,
        plant_id=plant_id,
        observed_at=observed_at,
        detector_version=detector_version,
        influence_factor=influence_factor,
        signal_type=signal_type,
        detector_role=detector_role,
    )


@dataclass(frozen=True)
class AnomalyAlert:
    scope: str
    event_id: str
    run_id: str
    plant_id: str
    plant_name: str
    region: str
    energy_source: str
    observed_at: datetime
    severity: Severity
    influence_factor: str
    summary: str
    detector_version: str
    detector_role: str
    audience: NotificationAudience
    route_key: str
    score: float | None = None
    actual_mwh: float | None = None
    predicted_mwh: float | None = None
    threshold_mwh: float | None = None
    evidence: Mapping[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        for name in (
            "event_id",
            "run_id",
            "plant_id",
            "plant_name",
            "region",
            "summary",
            "detector_version",
            "detector_role",
            "route_key",
            "scope",
            "energy_source",
        ):
            object.__setattr__(self, name, _required_text(name, getattr(self, name)))
        object.__setattr__(self, "observed_at", ensure_utc(self.observed_at))
        object.__setattr__(
            self,
            "severity",
            self.severity if isinstance(self.severity, Severity) else Severity(self.severity),
        )
        object.__setattr__(
            self,
            "audience",
            self.audience
            if isinstance(self.audience, NotificationAudience)
            else NotificationAudience(self.audience),
        )
        object.__setattr__(
            self,
            "influence_factor",
            validate_influence_factor(self.influence_factor),
        )
        for name in ("score", "actual_mwh", "predicted_mwh", "threshold_mwh"):
            value = getattr(self, name)
            if value is not None and not math.isfinite(float(value)):
                raise ValueError(f"{name} must be finite when provided")
        if self.score is not None and float(self.score) < 0:
            raise ValueError("score cannot be negative")
        for name in ("actual_mwh", "predicted_mwh"):
            value = getattr(self, name)
            if value is not None and float(value) < 0:
                raise ValueError(f"{name} cannot be negative")
        if self.threshold_mwh is not None and float(self.threshold_mwh) <= 0:
            raise ValueError("threshold_mwh must be positive")
        object.__setattr__(self, "evidence", dict(self.evidence))
        if self.scope != "operational":
            raise ValueError("notification alerts require scope=operational")
        if self.energy_source.strip().lower() != "solar":
            raise ValueError("plant-manager notifications require energy_source=solar")
        if any(
            phrase in self.summary
            for phrase in ("고장 확정", "센서 고장", "인버터 고장", "설비 고장")
        ):
            raise ValueError("notification summary cannot assert an unverified equipment failure")
        if (
            self.influence_factor == "data_quality_issue"
            and self.audience != NotificationAudience.DATA_OPERATOR
        ):
            raise ValueError("data-quality events must route to audience=data_operator")
        if self.detector_role not in {"deployed", "consensus"}:
            raise ValueError("detector_role must be deployed or consensus for alerts")

    @classmethod
    def from_operational_event(
        cls,
        event: Mapping[str, Any],
        *,
        detector_version: str | None = None,
        severity: Severity | str | None = None,
    ) -> "AnomalyAlert":
        """Map a full operational anomaly event into the alert contract.

        Dashboard ``prediction_signals`` are a historical Test top-N projection
        and intentionally fail this contract. Live alerting must consume the
        dedicated full event stream produced after operational inference.
        """

        contract = event.get("event_contract", event.get("schema_version"))
        if contract != OPERATIONAL_EVENT_CONTRACT:
            raise ValueError(
                f"event_contract must be {OPERATIONAL_EVENT_CONTRACT}"
            )
        if event.get("scope") != "operational":
            raise ValueError("only scope=operational anomaly events may be notified")
        if event.get("interpretation_limit") != INTERPRETATION_LIMIT:
            raise ValueError("operational event interpretation_limit is missing or incompatible")
        if not event.get("detected_at"):
            raise ValueError("operational anomaly event requires detected_at")
        observed_value = event.get("observed_at", event.get("timestamp"))
        if not observed_value:
            raise ValueError("operational anomaly event requires observed_at")
        observed_at = datetime.fromisoformat(str(observed_value).replace("Z", "+00:00"))
        observed_at = ensure_utc(observed_at)
        detected_at = ensure_utc(
            datetime.fromisoformat(str(event["detected_at"]).replace("Z", "+00:00"))
        )
        if detected_at < observed_at:
            raise ValueError("detected_at cannot be earlier than the observed timestamp")
        evidence = event.get("evidence")
        evidence = evidence if isinstance(evidence, Mapping) else {}
        detector = event.get("detector")
        detector = detector if isinstance(detector, Mapping) else {}
        actual = float(evidence.get("actual_mwh", event.get("y_true")))
        predicted = float(evidence.get("predicted_mwh", event.get("y_pred")))
        threshold = float(evidence.get("threshold_mwh", event.get("threshold")))
        ratio = float(evidence.get("exceedance_ratio", event.get("exceedance_ratio", 1.0)))
        threshold_unit = str(
            evidence.get("threshold_unit", event.get("threshold_unit")) or ""
        ).strip()
        if threshold_unit.lower() != "mwh":
            raise ValueError("operational notification energy values must use MWh")
        if actual < 0 or predicted < 0:
            raise ValueError("operational actual/predicted MWh values cannot be negative")
        if threshold <= 0:
            raise ValueError("operational threshold_mwh must be positive")
        if ratio < 0:
            raise ValueError("operational exceedance_ratio cannot be negative")
        resolved_severity = severity or event.get("severity") or (
            Severity.CRITICAL
            if ratio >= 3.0
            else Severity.HIGH
            if ratio >= 2.0
            else Severity.REVIEW
        )
        factor = event.get("influence_factor") or (
            "lower_than_expected" if actual < predicted else "higher_than_expected"
        )
        factor = validate_influence_factor(str(factor))
        signal_type = _required_text("signal_type", event.get("signal_type"))
        detector_role = _required_text(
            "detector_role", event.get("detector_role", detector.get("role"))
        )
        model = _required_text("detector id", detector.get("id", event.get("model")))
        detector_version = _required_text(
            "detector_version",
            detector_version
            or detector.get("version")
            or event.get("detector_version"),
        )
        plant_id = _required_text("plant_id", event.get("plant_id"))
        event_id = stable_event_id(
            model=model,
            plant_id=plant_id,
            observed_at=observed_at,
            detector_version=detector_version,
            influence_factor=factor,
            signal_type=signal_type,
            detector_role=detector_role,
        )
        supplied_event_id = str(event.get("event_id") or "").strip()
        if supplied_event_id and supplied_event_id != event_id:
            raise ValueError("event_id does not match the deterministic operational event ID")
        summary = str(event.get("summary") or "").strip() or (
            f"실측 {actual:,.3f} MWh, 예측 {predicted:,.3f} MWh, "
            f"판단 임계치 {threshold:,.3f} MWh"
        )
        return cls(
            scope="operational",
            event_id=event_id,
            run_id=_required_text("run_id", event.get("run_id")),
            plant_id=plant_id,
            plant_name=str(event.get("plant") or plant_id),
            region=str(event.get("region") or "지역 미확인"),
            energy_source=_required_text("energy_source", event.get("energy_source")),
            observed_at=observed_at,
            severity=Severity(resolved_severity),
            influence_factor=factor,
            summary=summary,
            detector_version=detector_version,
            detector_role=detector_role,
            audience=NotificationAudience(
                _required_text("audience", event.get("audience"))
            ),
            route_key=_required_text("route_key", event.get("route_key")),
            score=ratio,
            actual_mwh=actual,
            predicted_mwh=predicted,
            threshold_mwh=threshold,
            evidence={
                "model": model,
                "signal_type": signal_type,
                "detected_at": detected_at.isoformat(),
                "threshold_source": evidence.get(
                    "threshold_source", event.get("threshold_source")
                ),
                "threshold_unit": threshold_unit,
            },
        )


@dataclass(frozen=True)
class RecipientRoute:
    """A manager-directory reference and its ordered delivery fallback chain."""

    route_key: str
    plant_id: str
    recipient_id: str
    audience: NotificationAudience
    channels: tuple[NotificationChannel, ...] = (
        NotificationChannel.KAKAO_ALIMTALK,
        NotificationChannel.SMS,
    )

    def __post_init__(self) -> None:
        object.__setattr__(self, "route_key", _required_text("route_key", self.route_key))
        object.__setattr__(self, "plant_id", _required_text("plant_id", self.plant_id))
        object.__setattr__(self, "recipient_id", _required_text("recipient_id", self.recipient_id))
        object.__setattr__(
            self,
            "audience",
            self.audience
            if isinstance(self.audience, NotificationAudience)
            else NotificationAudience(self.audience),
        )
        if (
            self.audience == NotificationAudience.PLANT_MANAGER
            and self.plant_id == "*"
        ):
            raise ValueError("plant-manager routes must bind to one explicit plant_id")
        channels = tuple(
            item if isinstance(item, NotificationChannel) else NotificationChannel(item)
            for item in self.channels
        )
        if not channels:
            raise ValueError("at least one notification channel is required")
        if len(set(channels)) != len(channels):
            raise ValueError("notification fallback channels must be unique")
        object.__setattr__(self, "channels", channels)


@dataclass(frozen=True)
class NotificationEnvelope:
    event_id: str
    plant_id: str
    recipient_id: str
    channels: tuple[NotificationChannel, ...]
    template_id: str
    subject: str
    body: str
    metadata: Mapping[str, Any]
    idempotency_key: str

    @classmethod
    def for_anomaly(
        cls,
        alert: AnomalyAlert,
        route: RecipientRoute,
        *,
        template_id: str,
    ) -> "NotificationEnvelope":
        template_id = _required_text("template_id", template_id)
        severity_label = {
            Severity.REVIEW: "검토",
            Severity.HIGH: "주의",
            Severity.CRITICAL: "긴급",
        }[alert.severity]
        subject = f"[태양광 이상 신호/{severity_label}] {alert.plant_name}"
        factor_label = _FACTOR_LABELS[alert.influence_factor]
        observed = alert.observed_at.astimezone(ZoneInfo("Asia/Seoul")).strftime(
            "%Y-%m-%d %H:%M KST"
        )
        body = (
            f"{alert.region} {alert.plant_name}에서 검토가 필요한 이상 신호가 탐지되었습니다. "
            f"관측시각: {observed}. 판단 근거: {factor_label}; {alert.summary}. "
            f"{INTERPRETATION_LIMIT} 운영 화면에서 원시 계측과 현장 상태를 확인해 주세요."
        )
        key_payload = {
            "contract": NOTIFICATION_CONTRACT,
            "event_id": alert.event_id,
            "recipient_id": route.recipient_id,
            "template_id": template_id,
        }
        key = sha256(
            json.dumps(
                key_payload,
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":"),
            ).encode("utf-8")
        ).hexdigest()
        return cls(
            event_id=alert.event_id,
            plant_id=alert.plant_id,
            recipient_id=route.recipient_id,
            channels=route.channels,
            template_id=template_id,
            subject=subject,
            body=body,
            metadata={
                "contract": NOTIFICATION_CONTRACT,
                "run_id": alert.run_id,
                "energy_source": alert.energy_source,
                "severity": alert.severity.value,
                "influence_factor": alert.influence_factor,
                "influence_factor_label": factor_label,
                "observed_at": alert.observed_at.isoformat(),
                "observed_at_kst": alert.observed_at.astimezone(
                    ZoneInfo("Asia/Seoul")
                ).strftime("%Y-%m-%d %H:%M KST"),
                "detector_version": alert.detector_version,
                "detector_role": alert.detector_role,
                "audience": alert.audience.value,
                "route_key": alert.route_key,
                "score": alert.score,
                "summary": alert.summary,
                "plant_name": alert.plant_name,
                "region": alert.region,
                "evidence": dict(alert.evidence),
            },
            idempotency_key=key,
        )


@dataclass(frozen=True)
class DeliveryResult:
    provider_message_id: str
    accepted_at: datetime = field(default_factory=utc_now)

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "provider_message_id",
            _required_text("provider_message_id", self.provider_message_id),
        )
        object.__setattr__(self, "accepted_at", ensure_utc(self.accepted_at))


def normalize_channels(values: Sequence[str | NotificationChannel]) -> tuple[NotificationChannel, ...]:
    return tuple(
        item if isinstance(item, NotificationChannel) else NotificationChannel(item)
        for item in values
    )


def redact_sensitive(value: str) -> str:
    """Mask common Korean phone-number and bearer-token forms before audit writes."""

    text = str(value)
    text = re.sub(
        r"(?<!\d)(01\d)[- ]?(\d{3,4})[- ]?(\d{4})(?!\d)",
        lambda match: f"{match.group(1)}-****-{match.group(3)}",
        text,
    )
    text = re.sub(
        r"(?i)(bearer\s+)[A-Za-z0-9._~+\-/=]+",
        r"\1***",
        text,
    )
    text = re.sub(
        r"(?i)\b(api[_-]?key|api[_-]?secret|token|signature)\s*([:=])\s*[^,\s]+",
        lambda match: f"{match.group(1)}{match.group(2)}***",
        text,
    )
    return text[:1000]


def _required_text(name: str, value: Any) -> str:
    text = str(value).strip() if value is not None else ""
    if not text:
        raise ValueError(f"{name} cannot be blank")
    return text
