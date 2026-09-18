from __future__ import annotations

from dataclasses import replace
from datetime import datetime, timedelta, timezone
import hashlib
import hmac
import json
import sqlite3

import pytest

from solar_forecast.anomalies import INTERPRETATION_LIMIT
from solar_forecast.anomalies import verify_operational_event_batch
from solar_forecast.anomalies import write_operational_event_batch
from solar_forecast.cli import build_parser
from solar_forecast.cli import main
from solar_forecast.notifications import OPERATIONAL_EVENT_CONTRACT
from solar_forecast.notifications import AmbiguousNotificationError
from solar_forecast.notifications import AnomalyAlert
from solar_forecast.notifications import DeliveryResult
from solar_forecast.notifications import MappingContactDirectory
from solar_forecast.notifications import NotificationAudience
from solar_forecast.notifications import NotificationChannel
from solar_forecast.notifications import NotificationDispatcher
from solar_forecast.notifications import NotificationOutbox
from solar_forecast.notifications import NotificationService
from solar_forecast.notifications import NotificationSettings
from solar_forecast.notifications import NotificationStatus
from solar_forecast.notifications import PermanentNotificationError
from solar_forecast.notifications import ProviderMessage
from solar_forecast.notifications import RecipientRoute
from solar_forecast.notifications import RuntimeRouteDirectory
from solar_forecast.notifications import Severity
from solar_forecast.notifications import SolapiProvider
from solar_forecast.notifications import SolapiSettings
from solar_forecast.notifications import TransientNotificationError
from solar_forecast.notifications import stable_event_id


NOW = datetime(2026, 8, 31, 1, 0, tzinfo=timezone.utc)


def _event(**overrides):
    value = {
        "event_contract": OPERATIONAL_EVENT_CONTRACT,
        "scope": "operational",
        "run_id": "operational-20260831-01",
        "energy_source": "solar",
        "audience": "plant_manager",
        "route_key": "plant:plant-001:manager",
        "detector_role": "deployed",
        "signal_type": "residual_threshold_exceeded",
        "interpretation_limit": INTERPRETATION_LIMIT,
        "detected_at": "2026-08-31T01:05:00+00:00",
        "timestamp": "2026-08-31T01:00:00+00:00",
        "model": "hybrid",
        "plant_id": "plant-001",
        "plant": "테스트태양광",
        "region": "전북특별자치도",
        "y_true": 1.0,
        "y_pred": 5.0,
        "threshold": 2.0,
        "threshold_unit": "MWh",
        "threshold_source": "calibration",
        "exceedance_ratio": 2.0,
    }
    value.update(overrides)
    return value


def _alert(**overrides) -> AnomalyAlert:
    return AnomalyAlert.from_operational_event(
        _event(**overrides), detector_version="hybrid-prod-2026.08"
    )


def _route(
    *channels: NotificationChannel,
    audience: NotificationAudience = NotificationAudience.PLANT_MANAGER,
    plant_id: str = "plant-001",
) -> RecipientRoute:
    return RecipientRoute(
        route_key="plant:plant-001:manager",
        plant_id=plant_id,
        recipient_id="manager:plant-001:primary",
        audience=audience,
        channels=channels
        or (NotificationChannel.KAKAO_ALIMTALK, NotificationChannel.SMS),
    )


class _Directory:
    def __init__(self):
        self.calls = []

    def resolve(self, recipient_id, channel):
        self.calls.append((recipient_id, channel))
        return "010-1234-5678"


class _Provider:
    def __init__(self, channel, outcomes=None):
        self.channel = channel
        self.outcomes = list(outcomes or [DeliveryResult("provider-1", NOW)])
        self.messages = []

    def send(self, message):
        self.messages.append(message)
        outcome = self.outcomes.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return outcome


def _queue(tmp_path, alert=None, route=None, *, now=NOW):
    outbox = NotificationOutbox(tmp_path / "alerts.db")
    service = NotificationService(outbox, template_id="solar-anomaly-v1")
    result = service.enqueue_anomaly(alert or _alert(), [route or _route()], now=now)[0]
    return outbox, result


def _live_settings(tmp_path, **overrides):
    values = {
        "enabled": True,
        "dry_run": False,
        "live_confirmation": True,
        "outbox_path": tmp_path / "alerts.db",
        "max_attempts_per_channel": 2,
        "retry_base_seconds": 10,
        "retry_max_seconds": 60,
        "lease_seconds": 30,
        "batch_size": 10,
    }
    values.update(overrides)
    return NotificationSettings(**values)


def test_external_delivery_requires_environment_and_separate_live_gate(tmp_path) -> None:
    defaults = NotificationSettings.from_env({})
    assert defaults.enabled is False
    assert defaults.dry_run is True
    assert defaults.outbox_path.as_posix() == "artifacts/notifications/outbox.sqlite3"
    assert defaults.external_delivery_allowed is False

    configured = NotificationSettings.from_env(
        {"SOLAR_NOTIFY_ENABLED": "true", "SOLAR_NOTIFY_DRY_RUN": "false"}
    )
    assert configured.external_delivery_allowed is False
    assert configured.confirm_live(confirmed=True).external_delivery_allowed is True


@pytest.mark.parametrize(
    "overrides, message",
    [
        ({"scope": "evaluation"}, "scope=operational"),
        ({"event_contract": "dashboard-projection.v1"}, "event_contract"),
        ({"detected_at": None}, "detected_at"),
        ({"detector_role": "candidate"}, "detector_role"),
        ({"interpretation_limit": "설비 고장 확정"}, "interpretation_limit"),
    ],
)
def test_dashboard_evaluation_and_unreviewed_events_are_rejected(overrides, message) -> None:
    with pytest.raises(ValueError, match=message):
        _alert(**overrides)


def test_stable_event_id_distinguishes_multiple_signals_at_same_timestamp() -> None:
    base = {
        "model": "hybrid",
        "plant_id": "plant-001",
        "observed_at": NOW,
        "detector_version": "v1",
        "detector_role": "deployed",
    }
    low = stable_event_id(
        **base,
        influence_factor="lower_than_expected",
        signal_type="residual",
    )
    rapid = stable_event_id(
        **base,
        influence_factor="rapid_output_change",
        signal_type="ramp",
    )
    assert low != rapid


def test_data_quality_alert_cannot_be_routed_to_plant_manager() -> None:
    with pytest.raises(ValueError, match="data_operator"):
        AnomalyAlert(
            scope="operational",
            event_id="event-data-quality",
            run_id="quality-20260831",
            plant_id="plant-001",
            plant_name="테스트태양광",
            region="전북특별자치도",
            energy_source="solar",
            observed_at=NOW,
            severity=Severity.HIGH,
            influence_factor="data_quality_issue",
            summary="센서 수집 공백",
            detector_version="quality-v1",
            detector_role="deployed",
            audience=NotificationAudience.PLANT_MANAGER,
            route_key="plant:plant-001:manager",
        )


def test_plant_manager_route_cannot_cross_a_plant_boundary(tmp_path) -> None:
    outbox = NotificationOutbox(tmp_path / "alerts.db")
    service = NotificationService(outbox, template_id="solar-anomaly-v1")
    with pytest.raises(ValueError, match="plant_id"):
        service.enqueue_anomaly(_alert(), [_route(plant_id="plant-999")], now=NOW)
    assert outbox.status_counts() == {}


def test_enqueue_is_idempotent_and_persists_no_phone_or_api_secret(tmp_path) -> None:
    outbox, first = _queue(tmp_path)
    service = NotificationService(outbox, template_id="solar-anomaly-v1")
    second = service.enqueue_anomaly(_alert(), [_route()], now=NOW)[0]
    assert first.created is True
    assert second.created is False
    assert second.notification_id == first.notification_id

    database_text = (tmp_path / "alerts.db").read_bytes()
    assert b"01012345678" not in database_text
    assert b"super-secret-api-key" not in database_text
    with sqlite3.connect(tmp_path / "alerts.db") as connection:
        keys = connection.execute(
            "SELECT event_id, recipient_id, channel FROM delivery_keys ORDER BY channel"
        ).fetchall()
    assert keys == [
        (_alert().event_id, "manager:plant-001:primary", "kakao_alimtalk"),
        (_alert().event_id, "manager:plant-001:primary", "sms"),
    ]


def test_dry_run_suppresses_without_resolving_contact_or_calling_provider(tmp_path) -> None:
    outbox, queued = _queue(tmp_path)
    directory = _Directory()
    provider = _Provider(NotificationChannel.KAKAO_ALIMTALK)
    summary = NotificationDispatcher(
        NotificationSettings(outbox_path=tmp_path / "alerts.db"),
        outbox,
        directory,
        {NotificationChannel.KAKAO_ALIMTALK: provider},
    ).run_once(now=NOW)
    assert summary.candidates == 1
    assert summary.suppressed == 1
    assert directory.calls == []
    assert provider.messages == []
    assert outbox.get(queued.notification_id).status == NotificationStatus.DRY_RUN


def test_live_dispatch_records_provider_acceptance_and_not_delivery(tmp_path) -> None:
    outbox, queued = _queue(tmp_path, route=_route(NotificationChannel.SMS))
    directory = _Directory()
    provider = _Provider(NotificationChannel.SMS)
    dispatcher = NotificationDispatcher(
        _live_settings(tmp_path),
        outbox,
        directory,
        {NotificationChannel.SMS: provider},
    )
    summary = dispatcher.run_once(now=NOW)
    assert summary.accepted == 1
    record = outbox.get(queued.notification_id)
    assert record.status == NotificationStatus.ACCEPTED
    assert record.provider_message_id == "provider-1"
    assert "KST" in record.body
    assert INTERPRETATION_LIMIT in record.body
    assert outbox.audit_events(queued.notification_id)[-1]["action"] == "provider_accepted"
    assert dispatcher.run_once(now=NOW + timedelta(minutes=1)).candidates == 0
    assert len(provider.messages) == 1


def test_transient_failure_uses_exponential_retry_then_accepts(tmp_path) -> None:
    outbox, queued = _queue(tmp_path, route=_route(NotificationChannel.SMS))
    provider = _Provider(
        NotificationChannel.SMS,
        [TransientNotificationError("temporary"), DeliveryResult("provider-2", NOW)],
    )
    dispatcher = NotificationDispatcher(
        _live_settings(tmp_path),
        outbox,
        _Directory(),
        {NotificationChannel.SMS: provider},
    )
    first = dispatcher.run_once(now=NOW)
    assert first.retry_required == 1
    assert outbox.get(queued.notification_id).available_at == NOW + timedelta(seconds=10)
    assert dispatcher.run_once(now=NOW + timedelta(seconds=9)).candidates == 0
    second = dispatcher.run_once(now=NOW + timedelta(seconds=10))
    assert second.accepted == 1
    assert outbox.get(queued.notification_id).status == NotificationStatus.ACCEPTED


def test_ambiguous_provider_outcome_is_quarantined_without_automatic_retry(
    tmp_path,
) -> None:
    outbox, queued = _queue(tmp_path, route=_route(NotificationChannel.SMS))
    provider = _Provider(
        NotificationChannel.SMS,
        [AmbiguousNotificationError("provider response was lost")],
    )
    dispatcher = NotificationDispatcher(
        _live_settings(tmp_path),
        outbox,
        _Directory(),
        {NotificationChannel.SMS: provider},
    )
    first = dispatcher.run_once(now=NOW)
    assert first.in_doubt == 1
    assert first.retry_required == 0
    record = outbox.get(queued.notification_id)
    assert record.status == NotificationStatus.IN_DOUBT
    assert dispatcher.run_once(now=NOW + timedelta(days=1)).candidates == 0
    assert len(provider.messages) == 1
    assert outbox.audit_events(queued.notification_id)[-1]["action"] == (
        "provider_outcome_unknown"
    )


def test_permanent_kakao_failure_falls_back_to_sms_with_distinct_key(tmp_path) -> None:
    outbox, queued = _queue(tmp_path)
    kakao = _Provider(
        NotificationChannel.KAKAO_ALIMTALK,
        [PermanentNotificationError("template rejected")],
    )
    sms = _Provider(NotificationChannel.SMS)
    dispatcher = NotificationDispatcher(
        _live_settings(tmp_path),
        outbox,
        _Directory(),
        {
            NotificationChannel.KAKAO_ALIMTALK: kakao,
            NotificationChannel.SMS: sms,
        },
    )
    first = dispatcher.run_once(now=NOW)
    assert first.fallback_scheduled == 1
    sms_record = outbox.get(queued.notification_id)
    assert sms_record.channel == NotificationChannel.SMS
    assert sms_record.idempotency_key != kakao.messages[0].idempotency_key
    second = dispatcher.run_once(now=NOW)
    assert second.accepted == 1
    assert sms.messages[0].channel == NotificationChannel.SMS


def test_expired_dispatch_lease_is_in_doubt_not_automatically_resent(tmp_path) -> None:
    outbox, queued = _queue(tmp_path, route=_route(NotificationChannel.SMS))
    claim = outbox.claim_due(limit=1, lease_seconds=30, now=NOW)[0]
    assert claim.status == NotificationStatus.IN_PROGRESS
    assert outbox.claim_due(limit=1, lease_seconds=30, now=NOW + timedelta(seconds=31)) == []
    assert outbox.get(queued.notification_id).status == NotificationStatus.IN_DOUBT

    outbox.resolve_in_doubt(
        queued.notification_id,
        resolution_note="SOLAPI 콘솔에서 미접수 확인 후 재시도 승인",
        retry=True,
        now=NOW + timedelta(seconds=32),
    )
    reclaimed = outbox.claim_due(
        limit=1, lease_seconds=30, now=NOW + timedelta(seconds=32)
    )
    assert len(reclaimed) == 1
    assert reclaimed[0].attempts_total == 2


def test_failure_audit_redacts_phone_and_bearer_token(tmp_path) -> None:
    outbox, queued = _queue(tmp_path, route=_route(NotificationChannel.SMS))
    claim = outbox.claim_due(limit=1, lease_seconds=30, now=NOW)[0]
    outbox.mark_failed(
        queued.notification_id,
        expected_attempt=claim.attempts_total,
        error=(
            "to=010-1234-5678 Authorization: Bearer secret-token "
            "apiKey=runtime-key signature=runtime-signature"
        ),
        transient=False,
        max_attempts_per_channel=1,
        retry_base_seconds=10,
        retry_max_seconds=60,
        now=NOW,
    )
    serialized = json.dumps(outbox.audit_events(queued.notification_id), ensure_ascii=False)
    assert "010-****-5678" in serialized
    assert "010-1234-5678" not in serialized
    assert "secret-token" not in serialized
    assert "runtime-key" not in serialized
    assert "runtime-signature" not in serialized


class _Response:
    status_code = 200

    def json(self):
        return {"groupInfo": {"groupId": "G-test"}, "failedMessageList": []}


class _Session:
    def __init__(self):
        self.calls = []

    def post(self, endpoint, **kwargs):
        self.calls.append((endpoint, kwargs))
        return _Response()


def _solapi_settings() -> SolapiSettings:
    return SolapiSettings(
        api_key="runtime-key",
        api_secret="runtime-secret",
        sender_number="02-1234-5678",
        kakao_pf_id="runtime-pf-id",
        kakao_template_id="approved-template-id",
    )


def _provider_message(channel: NotificationChannel) -> ProviderMessage:
    alert = _alert()
    route = _route(channel)
    from solar_forecast.notifications.contracts import NotificationEnvelope

    envelope = NotificationEnvelope.for_anomaly(
        alert, route, template_id="solar-anomaly-v1"
    )
    return ProviderMessage(
        event_id=envelope.event_id,
        plant_id=envelope.plant_id,
        recipient_id=envelope.recipient_id,
        recipient_address="010-1234-5678",
        channel=channel,
        template_id=envelope.template_id,
        subject=envelope.subject,
        body=envelope.body,
        metadata=envelope.metadata,
        idempotency_key="delivery-key",
    )


def test_solapi_uses_v4_hmac_and_approved_alimtalk_template_contract() -> None:
    settings = _solapi_settings()
    fixed = datetime(2026, 8, 31, 1, 2, 3, tzinfo=timezone.utc)
    authorization = settings.authorization_header(now=fixed, salt="fixed-salt")
    date = "2026-08-31T01:02:03.000Z"
    expected = hmac.new(
        b"runtime-secret", f"{date}fixed-salt".encode(), hashlib.sha256
    ).hexdigest()
    assert authorization == (
        f"HMAC-SHA256 apiKey=runtime-key, date={date}, salt=fixed-salt, "
        f"signature={expected}"
    )
    assert "runtime-secret" not in repr(settings)
    assert "runtime-secret" not in json.dumps(settings.describe())

    session = _Session()
    result = SolapiProvider(
        NotificationChannel.KAKAO_ALIMTALK, settings, session=session
    ).send(_provider_message(NotificationChannel.KAKAO_ALIMTALK))
    assert result.provider_message_id == "G-test"
    endpoint, request = session.calls[0]
    assert endpoint.endswith("/messages/v4/send-many/detail")
    message = request["json"]["messages"][0]
    assert message["type"] == "ATA"
    assert message["to"] == "01012345678"
    assert message["kakaoOptions"]["pfId"] == "runtime-pf-id"
    assert message["kakaoOptions"]["templateId"] == "approved-template-id"
    assert message["kakaoOptions"]["disableSms"] is False
    assert "#{발전소명}" in message["kakaoOptions"]["variables"]
    assert message["kakaoOptions"]["variables"]["#{관측시각}"].endswith("KST")
    assert message["kakaoOptions"]["variables"]["#{판단근거}"] == "예측 대비 발전량 저하"
    assert request["headers"]["Authorization"].startswith("HMAC-SHA256 ")
    assert "Idempotency-Key" not in request["headers"]


def test_solapi_credentials_cannot_be_redirected_to_another_host() -> None:
    with pytest.raises(ValueError, match="official"):
        SolapiSettings(
            api_key="runtime-key",
            api_secret="runtime-secret",
            sender_number="02-1234-5678",
            kakao_pf_id="runtime-pf-id",
            kakao_template_id="approved-template-id",
            endpoint="https://example.test/messages/v4/send-many/detail",
        )


def test_solapi_long_sms_fallback_is_sent_as_lms_text_message() -> None:
    session = _Session()
    SolapiProvider(
        NotificationChannel.SMS, _solapi_settings(), session=session
    ).send(_provider_message(NotificationChannel.SMS))
    message = session.calls[0][1]["json"]["messages"][0]
    assert message["type"] == "LMS"
    assert message["text"]
    assert message["subject"]


@pytest.mark.parametrize("outcome", [TimeoutError("lost response"), 500, 408])
def test_solapi_ambiguous_transport_outcomes_are_not_classified_as_retryable(
    outcome,
) -> None:
    class Session:
        def post(self, *args, **kwargs):
            if isinstance(outcome, Exception):
                raise outcome
            return type(
                "Response",
                (),
                {"status_code": outcome, "json": lambda self: {}},
            )()

    with pytest.raises(AmbiguousNotificationError, match="outcome is unknown"):
        SolapiProvider(
            NotificationChannel.SMS,
            _solapi_settings(),
            session=Session(),
        ).send(_provider_message(NotificationChannel.SMS))


def test_solapi_accepts_registered_service_sender_but_requires_mobile_recipient() -> None:
    settings = SolapiSettings(
        api_key="runtime-key",
        api_secret="runtime-secret",
        sender_number="1577-1603",
        kakao_pf_id="runtime-pf-id",
        kakao_template_id="approved-template-id",
    )
    session = _Session()
    SolapiProvider(NotificationChannel.SMS, settings, session=session).send(
        _provider_message(NotificationChannel.SMS)
    )
    assert session.calls[0][1]["json"]["messages"][0]["from"] == "15771603"

    invalid_recipient = replace(
        _provider_message(NotificationChannel.SMS),
        recipient_address="02-1234-5678",
    )
    with pytest.raises(PermanentNotificationError, match="mobile number"):
        SolapiProvider(
            NotificationChannel.SMS,
            settings,
            session=_Session(),
        ).send(invalid_recipient)


def test_mapping_directory_is_runtime_only_and_missing_contact_falls_back() -> None:
    directory = MappingContactDirectory(
        {("manager-1", NotificationChannel.SMS): "01012345678"}
    )
    assert directory.resolve("manager-1", NotificationChannel.SMS) == "01012345678"
    with pytest.raises(PermanentNotificationError, match="No runtime contact"):
        directory.resolve("manager-1", NotificationChannel.KAKAO_ALIMTALK)


def test_operational_event_rejects_other_energy_sources_and_fault_claims() -> None:
    with pytest.raises(ValueError, match="energy_source=solar"):
        _alert(energy_source="wind")
    with pytest.raises(ValueError, match="equipment failure"):
        _alert(summary="인버터 고장 확정")


def test_operational_event_accepts_self_contained_nested_detector_and_evidence() -> None:
    event = _event()
    event["schema_version"] = event.pop("event_contract")
    event["observed_at"] = event.pop("timestamp")
    event["detector"] = {
        "id": event.pop("model"),
        "version": "hybrid-prod-2026.08",
        "role": event.pop("detector_role"),
    }
    event["evidence"] = {
        "actual_mwh": event.pop("y_true"),
        "predicted_mwh": event.pop("y_pred"),
        "threshold_mwh": event.pop("threshold"),
        "threshold_unit": event.pop("threshold_unit"),
        "threshold_source": event.pop("threshold_source"),
        "exceedance_ratio": event.pop("exceedance_ratio"),
    }
    alert = AnomalyAlert.from_operational_event(event)
    assert alert.run_id == "operational-20260831-01"
    assert alert.detector_version == "hybrid-prod-2026.08"
    assert alert.actual_mwh == 1.0


@pytest.mark.parametrize(
    "overrides, message",
    [
        ({"threshold_unit": "kWh"}, "must use MWh"),
        ({"y_true": -0.1}, "cannot be negative"),
        ({"y_pred": -0.1}, "cannot be negative"),
        ({"threshold": 0}, "must be positive"),
        ({"exceedance_ratio": -1}, "cannot be negative"),
    ],
)
def test_operational_notification_requires_clean_physical_units_and_ranges(
    overrides, message
) -> None:
    with pytest.raises(ValueError, match=message):
        _alert(**overrides)


def test_operational_event_batch_verifies_hash_count_and_excludes_contacts(tmp_path) -> None:
    batch = write_operational_event_batch(
        [_event()],
        tmp_path / "events",
        run_id="operational-20260831-01",
        detector_version="hybrid-prod-2026.08",
    )
    verified = verify_operational_event_batch(batch.events_path, batch.manifest_path)
    assert verified.event_count == 1
    assert list(verified.records())[0]["plant_id"] == "plant-001"

    batch.events_path.write_text(
        batch.events_path.read_text(encoding="utf-8") + " ", encoding="utf-8"
    )
    with pytest.raises(ValueError, match="SHA-256"):
        verify_operational_event_batch(batch.events_path, batch.manifest_path)

    with pytest.raises(ValueError, match="cannot contain contacts"):
        write_operational_event_batch(
            [{**_event(), "phone": "01012345678"}],
            tmp_path / "unsafe",
            run_id="operational-20260831-01",
            detector_version="v1",
        )


def _write_routes(path, *, include_phone=False) -> None:
    route = {
        "active": True,
        "route_key": "plant:plant-001:manager",
        "plant_id": "plant-001",
        "recipient_id": "manager:plant-001:primary",
        "audience": "plant_manager",
        "channels": ["kakao_alimtalk", "sms"],
        "phone_env": "PLANT_001_MANAGER_PHONE",
    }
    if include_phone:
        route["phone"] = "01012345678"
    path.write_text(
        json.dumps(
            {"schema_version": "solar-notification-routes.v1", "routes": [route]}
        ),
        encoding="utf-8",
    )


def test_runtime_route_directory_resolves_contacts_only_for_live_dispatch(tmp_path) -> None:
    path = tmp_path / "routes.json"
    _write_routes(path)
    preview = RuntimeRouteDirectory.from_file(
        path, environ={}, require_contacts=False
    )
    assert len(
        preview.routes_for(
            "plant:plant-001:manager",
            NotificationAudience.PLANT_MANAGER,
            "plant-001",
        )
    ) == 1
    with pytest.raises(ValueError, match="plant_id=plant-999"):
        preview.routes_for(
            "plant:plant-001:manager",
            NotificationAudience.PLANT_MANAGER,
            "plant-999",
        )
    with pytest.raises(ValueError, match="PLANT_001_MANAGER_PHONE"):
        RuntimeRouteDirectory.from_file(path, environ={}, require_contacts=True)
    live = RuntimeRouteDirectory.from_file(
        path,
        environ={"PLANT_001_MANAGER_PHONE": "01012345678"},
        require_contacts=True,
    )
    assert live.contacts.resolve(
        "manager:plant-001:primary", NotificationChannel.SMS
    ) == "01012345678"

    unsafe = tmp_path / "unsafe-routes.json"
    _write_routes(unsafe, include_phone=True)
    with pytest.raises(ValueError, match="phone numbers"):
        RuntimeRouteDirectory.from_file(
            unsafe, environ={}, require_contacts=False
        )


def test_notify_anomalies_cli_defaults_to_preview_and_has_explicit_live_gate() -> None:
    args = build_parser().parse_args(
        ["notify-anomalies", "--events", "events.jsonl", "--routes", "routes.json"]
    )
    assert args.live is False
    assert args.manifest is None
    assert args.template_contract == "solar-anomaly-v1"


def test_notify_anomalies_cli_dry_run_validates_manifest_and_uses_separate_outbox(
    tmp_path, capsys, monkeypatch
) -> None:
    batch = write_operational_event_batch(
        [_event()],
        tmp_path / "events",
        run_id="operational-20260831-01",
        detector_version="hybrid-prod-2026.08",
    )
    routes = tmp_path / "routes.json"
    _write_routes(routes)
    monkeypatch.setattr("solar_forecast.cli.notification_commands.PROJECT_ROOT", tmp_path)
    outbox = tmp_path / "artifacts" / "notifications" / "dry_run.sqlite3"
    main(
        [
            "notify-anomalies",
            "--events",
            str(batch.events_path),
            "--manifest",
            str(batch.manifest_path),
            "--routes",
            str(routes),
        ]
    )
    output = capsys.readouterr().out
    assert "Notification DRY-RUN" in output
    assert "1 queued" in output
    with sqlite3.connect(outbox) as connection:
        assert connection.execute(
            "SELECT status FROM notifications"
        ).fetchone()[0] == "dry_run"


def test_notify_anomalies_cli_validates_all_rows_before_mutating_outbox(
    tmp_path, monkeypatch
) -> None:
    invalid = _event(
        timestamp="2026-08-31T02:00:00+00:00",
        detected_at="2026-08-31T02:05:00+00:00",
        threshold_unit="kWh",
    )
    batch = write_operational_event_batch(
        [_event(), invalid],
        tmp_path / "events",
        run_id="operational-20260831-01",
        detector_version="hybrid-prod-2026.08",
    )
    routes = tmp_path / "routes.json"
    _write_routes(routes)
    monkeypatch.setattr("solar_forecast.cli.notification_commands.PROJECT_ROOT", tmp_path)
    outbox = tmp_path / "artifacts" / "notifications" / "dry_run.sqlite3"

    with pytest.raises(ValueError, match="must use MWh"):
        main(
            [
                "notify-anomalies",
                "--events",
                str(batch.events_path),
                "--manifest",
                str(batch.manifest_path),
                "--routes",
                str(routes),
            ]
        )
    assert not outbox.exists()


def test_notify_anomalies_cli_live_is_blocked_without_environment_gate(monkeypatch) -> None:
    monkeypatch.setenv("SOLAR_NOTIFY_ENABLED", "false")
    monkeypatch.setenv("SOLAR_NOTIFY_DRY_RUN", "false")
    with pytest.raises(SystemExit, match="Live notifications are blocked"):
        main(
            [
                "notify-anomalies",
                "--events",
                "not-read.jsonl",
                "--routes",
                "not-read.json",
                "--live",
            ]
        )


def test_notify_anomalies_cli_rejects_any_custom_preview_outbox(tmp_path) -> None:
    batch = write_operational_event_batch(
        [_event()],
        tmp_path / "events",
        run_id="operational-20260831-01",
        detector_version="hybrid-prod-2026.08",
    )
    routes = tmp_path / "routes.json"
    _write_routes(routes)
    with pytest.raises(SystemExit, match="does not accept --outbox"):
        main(
            [
                "notify-anomalies",
                "--events",
                str(batch.events_path),
                "--routes",
                str(routes),
                "--outbox",
                str(tmp_path / "preview.db"),
            ]
        )
