"""운영 이상치 이벤트 ID 생성, JSONL 무결성 manifest 기록과 검증."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from hashlib import sha256
import json
from pathlib import Path
from typing import Any, Iterable, Iterator, Mapping

from solar_forecast.infrastructure.artifact_store import (
    replace_file_atomic,
    sha256_file,
    write_json_atomic,
)
from solar_forecast.anomalies.influence_policy import validate_influence_factor


EVENT_CONTRACT = "solar-anomaly-event.v1"
EVENT_MANIFEST_CONTRACT = "solar-anomaly-event-manifest.v1"

_FORBIDDEN_CONTACT_KEYS = frozenset(
    {
        "phone",
        "phone_number",
        "mobile",
        "recipient",
        "recipient_id",
        "contact",
        "api_key",
        "api_secret",
        "token",
    }
)


@dataclass(frozen=True)
class OperationalEventBatch:
    """A verified, replayable operational anomaly-event stream."""

    events_path: Path
    manifest_path: Path
    run_id: str
    detector_version: str
    event_count: int
    sha256: str

    def records(self) -> Iterator[dict[str, Any]]:
        """Read one record at a time after the whole file passed integrity checks."""

        yield from _iter_records(
            self.events_path,
            expected_run_id=self.run_id,
            expected_detector_version=self.detector_version,
        )


def stable_operational_event_id(
    *,
    detector_id: str,
    plant_id: str,
    observed_at: datetime,
    detector_version: str,
    influence_factor: str,
    signal_type: str,
    detector_role: str,
) -> str:
    """Create the stable event key shared by producers and notification consumers."""

    if observed_at.tzinfo is None or observed_at.utcoffset() is None:
        raise ValueError("observed_at must include a timezone")
    payload = {
        "contract": EVENT_CONTRACT,
        "detector_id": _required_text("detector_id", detector_id),
        "plant_id": _required_text("plant_id", plant_id),
        "observed_at": observed_at.astimezone(timezone.utc).isoformat(),
        "detector_version": _required_text(
            "detector_version", detector_version
        ),
        "influence_factor": validate_influence_factor(influence_factor),
        "signal_type": _required_text("signal_type", signal_type),
        "detector_role": _required_text("detector_role", detector_role),
    }
    encoded = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return f"anomaly:{sha256(encoded).hexdigest()}"


def operational_event_id(record: Mapping[str, Any]) -> str:
    """Derive an event ID from either the flat or nested v1 event representation."""

    detector = record.get("detector")
    detector = detector if isinstance(detector, Mapping) else {}
    evidence = record.get("evidence")
    evidence = evidence if isinstance(evidence, Mapping) else {}
    observed_value = record.get("observed_at", record.get("timestamp"))
    if not observed_value:
        raise ValueError("operational anomaly event requires observed_at")
    observed_at = datetime.fromisoformat(str(observed_value).replace("Z", "+00:00"))
    if observed_at.tzinfo is None or observed_at.utcoffset() is None:
        raise ValueError("operational anomaly observed_at must include a timezone")
    factor = record.get("influence_factor")
    if not factor:
        actual = float(evidence.get("actual_mwh", record.get("y_true")))
        predicted = float(evidence.get("predicted_mwh", record.get("y_pred")))
        factor = (
            "lower_than_expected" if actual < predicted else "higher_than_expected"
        )
    return stable_operational_event_id(
        detector_id=_required_text(
            "detector id", detector.get("id", record.get("model"))
        ),
        plant_id=_required_text("plant_id", record.get("plant_id")),
        observed_at=observed_at,
        detector_version=_required_text(
            "detector_version",
            detector.get("version", record.get("detector_version")),
        ),
        influence_factor=str(factor),
        signal_type=_required_text("signal_type", record.get("signal_type")),
        detector_role=_required_text(
            "detector_role", record.get("detector_role", detector.get("role"))
        ),
    )


def write_operational_event_batch(
    records: Iterable[Mapping[str, Any]],
    output_dir: Path,
    *,
    run_id: str,
    detector_version: str,
) -> OperationalEventBatch:
    """Atomically write JSONL plus a count/hash manifest without buffering all events."""

    run_id = _required_text("run_id", run_id)
    detector_version = _required_text("detector_version", detector_version)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    events_path = output_dir / "events.jsonl"
    temporary = events_path.with_suffix(".jsonl.part")
    count = 0
    try:
        with temporary.open("w", encoding="utf-8", newline="\n") as target:
            for count, source in enumerate(records, start=1):
                record = dict(source)
                existing_run_id = str(record.get("run_id") or run_id).strip()
                detector = record.get("detector")
                detector = detector if isinstance(detector, Mapping) else {}
                existing_version = str(
                    detector.get("version")
                    or record.get("detector_version")
                    or detector_version
                ).strip()
                if existing_run_id != run_id:
                    raise ValueError(
                        f"event run_id does not match batch run_id at line {count}"
                    )
                if existing_version != detector_version:
                    raise ValueError(
                        "event detector_version does not match batch detector_version "
                        f"at line {count}"
                    )
                record["run_id"] = run_id
                record["detector_version"] = detector_version
                expected_event_id = operational_event_id(record)
                supplied_event_id = str(record.get("event_id") or "").strip()
                if supplied_event_id and supplied_event_id != expected_event_id:
                    raise ValueError(
                        f"event_id is not deterministic at line {count}"
                    )
                record["event_id"] = expected_event_id
                _validate_event_shell(record, line_number=count)
                target.write(
                    json.dumps(
                        record,
                        ensure_ascii=False,
                        sort_keys=True,
                        separators=(",", ":"),
                        allow_nan=False,
                    )
                    + "\n"
                )
    except Exception:
        temporary.unlink(missing_ok=True)
        raise
    replace_file_atomic(temporary, events_path)
    digest = sha256_file(events_path)
    manifest_path = output_dir / "events.manifest.json"
    write_json_atomic(
        manifest_path,
        {
            "manifest_contract": EVENT_MANIFEST_CONTRACT,
            "event_contract": EVENT_CONTRACT,
            "scope": "operational",
            "run_id": run_id,
            "detector_version": detector_version,
            "events_file": events_path.name,
            "event_count": count,
            "sha256": digest,
        },
    )
    return OperationalEventBatch(
        events_path=events_path,
        manifest_path=manifest_path,
        run_id=run_id,
        detector_version=detector_version,
        event_count=count,
        sha256=digest,
    )


def verify_operational_event_batch(
    events_path: Path,
    manifest_path: Path | None = None,
) -> OperationalEventBatch:
    """Fail closed before enqueueing if lineage, count, contract or JSONL is wrong."""

    events_path = Path(events_path)
    manifest_path = Path(manifest_path) if manifest_path else events_path.with_name(
        "events.manifest.json"
    )
    if not events_path.is_file():
        raise FileNotFoundError(f"Operational anomaly events are missing: {events_path}")
    if not manifest_path.is_file():
        raise FileNotFoundError(f"Operational anomaly manifest is missing: {manifest_path}")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if not isinstance(manifest, dict):
        raise ValueError("operational anomaly manifest must be a JSON object")
    if manifest.get("manifest_contract") != EVENT_MANIFEST_CONTRACT:
        raise ValueError(f"manifest_contract must be {EVENT_MANIFEST_CONTRACT}")
    if manifest.get("event_contract") != EVENT_CONTRACT:
        raise ValueError(f"event_contract must be {EVENT_CONTRACT}")
    if manifest.get("scope") != "operational":
        raise ValueError("only scope=operational event batches can be notified")
    if manifest.get("events_file") != events_path.name:
        raise ValueError("manifest events_file does not match the selected JSONL")
    run_id = _required_text("run_id", manifest.get("run_id"))
    detector_version = _required_text(
        "detector_version", manifest.get("detector_version")
    )
    expected_hash = _required_text("manifest sha256", manifest.get("sha256"))
    actual_hash = sha256_file(events_path)
    if actual_hash != expected_hash:
        raise ValueError("operational anomaly JSONL SHA-256 does not match its manifest")
    expected_count = int(manifest.get("event_count", -1))
    actual_count = sum(
        1
        for _ in _iter_records(
            events_path,
            expected_run_id=run_id,
            expected_detector_version=detector_version,
        )
    )
    if expected_count < 0 or actual_count != expected_count:
        raise ValueError(
            "operational anomaly event count does not match its manifest: "
            f"expected={expected_count}, actual={actual_count}"
        )
    return OperationalEventBatch(
        events_path=events_path,
        manifest_path=manifest_path,
        run_id=run_id,
        detector_version=detector_version,
        event_count=actual_count,
        sha256=actual_hash,
    )


def _iter_records(
    path: Path,
    *,
    expected_run_id: str | None = None,
    expected_detector_version: str | None = None,
) -> Iterator[dict[str, Any]]:
    with Path(path).open("r", encoding="utf-8") as source:
        for line_number, raw_line in enumerate(source, start=1):
            if not raw_line.strip():
                raise ValueError(f"blank line in operational event JSONL at line {line_number}")
            try:
                record = json.loads(raw_line)
            except json.JSONDecodeError as exc:
                raise ValueError(
                    f"invalid operational event JSON at line {line_number}: {exc.msg}"
                ) from exc
            if not isinstance(record, dict):
                raise ValueError(
                    f"operational event at line {line_number} must be a JSON object"
                )
            _validate_event_shell(record, line_number=line_number)
            if expected_run_id is not None and record.get("run_id") != expected_run_id:
                raise ValueError(
                    f"event run_id does not match its manifest at line {line_number}"
                )
            version = record.get("detector_version")
            detector = record.get("detector")
            if isinstance(detector, Mapping):
                version = detector.get("version", version)
            if (
                expected_detector_version is not None
                and version != expected_detector_version
            ):
                raise ValueError(
                    "event detector_version does not match its manifest at line "
                    f"{line_number}"
                )
            yield record


def _validate_event_shell(record: Mapping[str, Any], *, line_number: int) -> None:
    contract = record.get("event_contract", record.get("schema_version"))
    if contract != EVENT_CONTRACT:
        raise ValueError(
            f"event_contract must be {EVENT_CONTRACT} at line {line_number}"
        )
    if record.get("scope") != "operational":
        raise ValueError(
            f"only scope=operational events are accepted at line {line_number}"
        )
    supplied_event_id = str(record.get("event_id") or "").strip()
    if not supplied_event_id:
        raise ValueError(f"operational event_id is required at line {line_number}")
    if supplied_event_id != operational_event_id(record):
        raise ValueError(
            f"operational event_id is not deterministic at line {line_number}"
        )
    forbidden = _find_forbidden_keys(record)
    if forbidden:
        names = ", ".join(sorted(forbidden))
        raise ValueError(
            f"operational events cannot contain contacts or credentials at line "
            f"{line_number}: {names}"
        )


def _find_forbidden_keys(value: Any) -> set[str]:
    found: set[str] = set()
    if isinstance(value, Mapping):
        for key, child in value.items():
            normalized = str(key).strip().lower()
            if normalized in _FORBIDDEN_CONTACT_KEYS:
                found.add(normalized)
            found.update(_find_forbidden_keys(child))
    elif isinstance(value, list):
        for child in value:
            found.update(_find_forbidden_keys(child))
    return found


def _required_text(name: str, value: Any) -> str:
    text = str(value).strip() if value is not None else ""
    if not text:
        raise ValueError(f"{name} cannot be blank")
    return text


__all__ = [
    "EVENT_CONTRACT",
    "EVENT_MANIFEST_CONTRACT",
    "OperationalEventBatch",
    "operational_event_id",
    "stable_operational_event_id",
    "verify_operational_event_batch",
    "write_operational_event_batch",
]
