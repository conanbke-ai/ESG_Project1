from __future__ import annotations

import argparse
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import ctypes
import json
import os
from pathlib import Path
import sqlite3
import subprocess
import sys
import time
from typing import Any, Iterable


@dataclass(frozen=True)
class CandidateSpec:
    index: int
    total: int
    horizon_hours: int
    model: str
    candidate_id: str
    study_base: str
    max_trials: int
    progress_unit: str
    progress_total: int | None


@dataclass(frozen=True)
class TrialStatus:
    study_name: str
    study_id: int
    counts: dict[str, int]
    latest_trial_id: int | None
    latest_number: int | None
    latest_state: str | None
    latest_started_at: str | None
    latest_completed_at: str | None
    latest_heartbeat_at: str | None
    latest_intermediate_step: int | None
    latest_intermediate_value: float | None
    latest_trial_best_step: int | None
    latest_trial_best_value: float | None
    latest_params: dict[str, Any]
    best_number: int | None
    best_value: float | None
    recent_trials: list[dict[str, Any]]


@dataclass(frozen=True)
class LockStatus:
    exists: bool
    pid: int | None
    model: str | None
    process_running: bool | None
    raw: str | None = None


def _load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _resolve(root: Path, value: str | Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else root / path


def _model_progress_contract(
    root: Path,
    model_config_path: str,
) -> tuple[str, int | None]:
    path = _resolve(root, model_config_path)
    values = _load_json(path)
    optimizer = values.get("optimizer", {})
    if values.get("model") == "cnn_bilstm":
        return "epoch", int(optimizer.get("trial_epochs", 0)) or None
    if values.get("model") == "xgboost":
        return "round", int(optimizer.get("trial_max_estimators", 0)) or None
    return "step", None


def build_candidate_specs(root: Path, config_path: Path) -> list[CandidateSpec]:
    values = _load_json(config_path)
    raw: list[tuple[int, str, str, str, int, str, int | None]] = []
    for horizon in values["horizons_hours"]:
        for model, settings in values["models"].items():
            unit, total_steps = _model_progress_contract(root, settings["config"])
            lengths = (
                settings.get("sequence_lengths", [1])
                if model == "cnn_bilstm"
                else [1]
            )
            for feature_set in settings["feature_sets"]:
                for length in lengths:
                    candidate_id = (
                        f"{feature_set}_lookback_{length}h"
                        if model == "cnn_bilstm"
                        else feature_set
                    )
                    study_base = f"historical_{model}_{horizon}h_{candidate_id}"
                    raw.append(
                        (
                            int(horizon),
                            model,
                            candidate_id,
                            study_base,
                            int(settings["max_trials_per_candidate"]),
                            unit,
                            total_steps,
                        )
                    )
    total = len(raw)
    return [
        CandidateSpec(
            index=i,
            total=total,
            horizon_hours=item[0],
            model=item[1],
            candidate_id=item[2],
            study_base=item[3],
            max_trials=item[4],
            progress_unit=item[5],
            progress_total=item[6],
        )
        for i, item in enumerate(raw, 1)
    ]


def _pid_running_posix(pid: int) -> bool:
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    return True


def _pid_running_windows(pid: int) -> bool:
    process_query_limited_information = 0x1000
    still_active = 259
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(
        process_query_limited_information,
        False,
        int(pid),
    )
    if not handle:
        return False
    try:
        code = ctypes.c_ulong()
        if not kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
            return True
        return int(code.value) == still_active
    finally:
        kernel32.CloseHandle(handle)


def pid_running(pid: int) -> bool:
    if pid <= 0:
        return False
    return (
        _pid_running_windows(pid)
        if os.name == "nt"
        else _pid_running_posix(pid)
    )


def read_lock(lock_path: Path) -> LockStatus:
    if not lock_path.exists():
        return LockStatus(False, None, None, None)
    raw = lock_path.read_text(encoding="utf-8", errors="replace")
    try:
        payload = json.loads(raw)
        pid = int(payload.get("pid")) if payload.get("pid") is not None else None
        model = (
            str(payload.get("model"))
            if payload.get("model") is not None
            else None
        )
    except (json.JSONDecodeError, TypeError, ValueError):
        return LockStatus(True, None, None, None, raw=raw)
    return LockStatus(
        True,
        pid,
        model,
        pid_running(pid) if pid is not None else None,
        raw=raw,
    )


def latest_benchmark_run(
    benchmarks_root: Path,
) -> tuple[Path | None, dict[str, Any] | None]:
    candidates: list[tuple[datetime, Path, dict[str, Any]]] = []
    if not benchmarks_root.exists():
        return None, None
    for manifest_path in benchmarks_root.glob("*/manifest.json"):
        try:
            manifest = _load_json(manifest_path)
        except (OSError, json.JSONDecodeError):
            continue
        stamp = manifest.get("created_at_utc")
        try:
            when = (
                datetime.fromisoformat(str(stamp).replace("Z", "+00:00"))
                if stamp
                else datetime.fromtimestamp(
                    manifest_path.stat().st_mtime,
                    timezone.utc,
                )
            )
        except ValueError:
            when = datetime.fromtimestamp(
                manifest_path.stat().st_mtime,
                timezone.utc,
            )
        if when.tzinfo is None:
            when = when.replace(tzinfo=timezone.utc)
        candidates.append((when, manifest_path.parent, manifest))
    if not candidates:
        return None, None
    _, run_dir, manifest = max(candidates, key=lambda item: item[0])
    return run_dir, manifest


def candidate_manifest_status(
    run_dir: Path | None,
    spec: CandidateSpec,
) -> tuple[str, Path | None, dict[str, Any] | None]:
    if run_dir is None:
        return "absent", None, None
    root = (
        run_dir
        / "candidates"
        / f"horizon_{spec.horizon_hours}h"
        / spec.model
        / spec.candidate_id
    )
    if not root.exists():
        return "absent", None, None
    manifests = sorted(
        root.glob("*/manifest.json"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    if not manifests:
        return "absent", None, None
    path = manifests[0]
    try:
        data = _load_json(path)
    except (OSError, json.JSONDecodeError):
        return "invalid", path, None
    return str(data.get("status", "unknown")), path, data


def _connect_readonly(db_path: Path) -> sqlite3.Connection:
    uri = db_path.resolve().as_uri() + "?mode=ro"
    connection = sqlite3.connect(uri, uri=True, timeout=2.0)
    connection.row_factory = sqlite3.Row
    return connection


def _study_rows(
    connection: sqlite3.Connection,
    study_base: str,
) -> list[sqlite3.Row]:
    return list(
        connection.execute(
            """
            SELECT study_id, study_name
            FROM studies
            WHERE study_name = ? OR study_name LIKE ?
            ORDER BY study_id DESC
            """,
            (study_base, f"{study_base}_%"),
        )
    )


def _latest_heartbeat(
    connection: sqlite3.Connection,
    trial_id: int,
) -> str | None:
    exists = connection.execute(
        """
        SELECT 1
        FROM sqlite_master
        WHERE type = 'table' AND name = 'trial_heartbeats'
        """
    ).fetchone()
    if exists is None:
        return None
    row = connection.execute(
        """
        SELECT heartbeat
        FROM trial_heartbeats
        WHERE trial_id = ?
        ORDER BY heartbeat DESC
        LIMIT 1
        """,
        (trial_id,),
    ).fetchone()
    return str(row["heartbeat"]) if row is not None else None


def _external_trial_params(
    connection: sqlite3.Connection,
    trial_id: int,
) -> dict[str, Any]:
    try:
        from optuna.distributions import json_to_distribution
    except ImportError:
        return {}

    params: dict[str, Any] = {}
    rows = connection.execute(
        """
        SELECT param_name, param_value, distribution_json
        FROM trial_params
        WHERE trial_id = ?
        ORDER BY param_id
        """,
        (trial_id,),
    )
    for row in rows:
        try:
            distribution = json_to_distribution(str(row["distribution_json"]))
            params[str(row["param_name"])] = distribution.to_external_repr(
                float(row["param_value"])
            )
        except Exception:
            params[str(row["param_name"])] = row["param_value"]
    return params


def _trial_status_for_study(
    connection: sqlite3.Connection,
    study_id: int,
    study_name: str,
) -> TrialStatus:
    counts = {
        str(row["state"]): int(row["n"])
        for row in connection.execute(
            """
            SELECT state, COUNT(*) AS n
            FROM trials
            WHERE study_id = ?
            GROUP BY state
            """,
            (study_id,),
        )
    }
    latest = connection.execute(
        """
        SELECT trial_id, number, state, datetime_start, datetime_complete
        FROM trials
        WHERE study_id = ?
        ORDER BY number DESC
        LIMIT 1
        """,
        (study_id,),
    ).fetchone()
    latest_step = latest_value = None
    latest_trial_best_step = latest_trial_best_value = None
    latest_heartbeat_at = None
    latest_params: dict[str, Any] = {}
    if latest is not None:
        latest_heartbeat_at = _latest_heartbeat(
            connection,
            int(latest["trial_id"]),
        )
        latest_params = _external_trial_params(
            connection,
            int(latest["trial_id"]),
        )
        intermediate = connection.execute(
            """
            SELECT step, intermediate_value
            FROM trial_intermediate_values
            WHERE trial_id = ?
            ORDER BY step DESC
            LIMIT 1
            """,
            (int(latest["trial_id"]),),
        ).fetchone()
        if intermediate is not None:
            latest_step = int(intermediate["step"])
            latest_value = float(intermediate["intermediate_value"])
        trial_best = connection.execute(
            """
            SELECT step, intermediate_value
            FROM trial_intermediate_values
            WHERE trial_id = ?
            ORDER BY intermediate_value ASC, step ASC
            LIMIT 1
            """,
            (int(latest["trial_id"]),),
        ).fetchone()
        if trial_best is not None:
            latest_trial_best_step = int(trial_best["step"])
            latest_trial_best_value = float(trial_best["intermediate_value"])
    best = connection.execute(
        """
        SELECT t.number, tv.value
        FROM trials t
        JOIN trial_values tv
          ON tv.trial_id = t.trial_id
         AND tv.objective = 0
        WHERE t.study_id = ?
          AND t.state = 'COMPLETE'
          AND tv.value_type = 'FINITE'
        ORDER BY tv.value ASC, t.number ASC
        LIMIT 1
        """,
        (study_id,),
    ).fetchone()
    recent_trials = [
        {
            "number": int(row["number"]),
            "state": str(row["state"]),
            "value": (
                float(row["display_value"])
                if row["display_value"] is not None
                else None
            ),
            "started_at": (
                str(row["datetime_start"])
                if row["datetime_start"] is not None
                else None
            ),
            "completed_at": (
                str(row["datetime_complete"])
                if row["datetime_complete"] is not None
                else None
            ),
            "last_step": (
                int(row["last_step"])
                if row["last_step"] is not None
                else None
            ),
        }
        for row in connection.execute(
            """
            SELECT
                t.number,
                t.state,
                t.datetime_start,
                t.datetime_complete,
                COALESCE(
                    tv.value,
                    (
                        SELECT tiv.intermediate_value
                        FROM trial_intermediate_values tiv
                        WHERE tiv.trial_id = t.trial_id
                        ORDER BY tiv.step DESC
                        LIMIT 1
                    )
                ) AS display_value,
                (
                    SELECT tiv.step
                    FROM trial_intermediate_values tiv
                    WHERE tiv.trial_id = t.trial_id
                    ORDER BY tiv.step DESC
                    LIMIT 1
                ) AS last_step
            FROM trials t
            LEFT JOIN trial_values tv
              ON tv.trial_id = t.trial_id
             AND tv.objective = 0
            WHERE t.study_id = ?
            ORDER BY t.number DESC
            LIMIT 5
            """,
            (study_id,),
        )
    ]
    return TrialStatus(
        study_name=study_name,
        study_id=study_id,
        counts=counts,
        latest_trial_id=int(latest["trial_id"]) if latest is not None else None,
        latest_number=int(latest["number"]) if latest is not None else None,
        latest_state=str(latest["state"]) if latest is not None else None,
        latest_started_at=(
            str(latest["datetime_start"])
            if latest is not None and latest["datetime_start"] is not None
            else None
        ),
        latest_completed_at=(
            str(latest["datetime_complete"])
            if latest is not None and latest["datetime_complete"] is not None
            else None
        ),
        latest_heartbeat_at=latest_heartbeat_at,
        latest_intermediate_step=latest_step,
        latest_intermediate_value=latest_value,
        latest_trial_best_step=latest_trial_best_step,
        latest_trial_best_value=latest_trial_best_value,
        latest_params=latest_params,
        best_number=int(best["number"]) if best is not None else None,
        best_value=float(best["value"]) if best is not None else None,
        recent_trials=recent_trials,
    )


def study_status(db_path: Path, study_base: str) -> TrialStatus | None:
    if not db_path.exists():
        return None
    connection = _connect_readonly(db_path)
    try:
        rows = _study_rows(connection, study_base)
        if not rows:
            return None
        scored: list[tuple[str, int, sqlite3.Row]] = []
        for row in rows:
            latest = connection.execute(
                """
                SELECT COALESCE(
                    MAX(COALESCE(datetime_complete, datetime_start)),
                    ''
                ) AS latest_time
                FROM trials
                WHERE study_id = ?
                """,
                (int(row["study_id"]),),
            ).fetchone()
            scored.append(
                (
                    str(latest["latest_time"]),
                    int(row["study_id"]),
                    row,
                )
            )
        _, _, chosen = max(scored, key=lambda item: (item[0], item[1]))
        return _trial_status_for_study(
            connection,
            int(chosen["study_id"]),
            str(chosen["study_name"]),
        )
    finally:
        connection.close()


def summarize(
    root: Path,
    config_path: Path,
    db_path: Path,
    benchmarks_root: Path,
    lock_path: Path,
) -> dict[str, Any]:
    specs = build_candidate_specs(root, config_path)
    lock = read_lock(lock_path)
    run_dir, run_manifest = latest_benchmark_run(benchmarks_root)
    candidate_rows: list[dict[str, Any]] = []
    active_index: int | None = None
    for spec in specs:
        manifest_status, manifest_path, _ = candidate_manifest_status(
            run_dir,
            spec,
        )
        try:
            optuna = study_status(db_path, spec.study_base)
        except sqlite3.OperationalError as exc:
            optuna = None
            db_error = f"{type(exc).__name__}: {exc}"
        else:
            db_error = None
        row = {
            **asdict(spec),
            "candidate_manifest_status": manifest_status,
            "candidate_manifest": (
                str(manifest_path)
                if manifest_path
                else None
            ),
            "optuna": asdict(optuna) if optuna else None,
            "db_error": db_error,
        }
        candidate_rows.append(row)
        if active_index is None and manifest_status == "running":
            active_index = spec.index
    if active_index is None:
        for row in candidate_rows:
            optuna = row["optuna"]
            if optuna and optuna.get("latest_state") == "RUNNING":
                active_index = int(row["index"])
                break

    run_status = str(run_manifest.get("status")) if run_manifest else None
    if run_status == "completed":
        overall = "COMPLETED"
    elif run_status == "failed":
        overall = "FAILED"
    elif (
        active_index is not None
        and lock.exists
        and lock.process_running is True
    ):
        overall = "RUNNING"
    elif run_status == "running" or active_index is not None:
        overall = "INTERRUPTED_OR_STALE"
    elif any(row["optuna"] for row in candidate_rows):
        overall = "OPTUNA_STATE_FOUND"
    else:
        overall = "NOT_STARTED"

    return {
        "overall_status": overall,
        "benchmark_run": str(run_dir) if run_dir else None,
        "benchmark_manifest_status": run_status,
        "benchmark_created_at": (
            run_manifest.get("created_at_utc") if run_manifest else None
        ),
        "lock": asdict(lock),
        "active_candidate_index": active_index,
        "candidates": candidate_rows,
        "db_path": str(db_path),
        "config_path": str(config_path),
    }


def _fmt_value(value: float | None) -> str:
    return "-" if value is None else f"{value:.6f} MWh"


def print_human(summary: dict[str, Any]) -> None:
    print(f"상태: {summary['overall_status']}")
    print(f"benchmark: {summary['benchmark_run'] or '-'}")
    lock = summary["lock"]
    if lock["exists"]:
        live = (
            "실행 중"
            if lock["process_running"] is True
            else "stale/확인 불가"
        )
        print(
            f"lock: pid={lock['pid']} "
            f"model={lock['model']} ({live})"
        )
    else:
        print("lock: 없음")
    print()

    for row in summary["candidates"]:
        optuna = row["optuna"]
        marker = (
            ">"
            if row["index"] == summary["active_candidate_index"]
            else " "
        )
        print(
            f"{marker} 후보 {row['index']}/{row['total']} | "
            f"{row['model']} | {row['horizon_hours']}h | "
            f"{row['candidate_id']}"
        )
        if row["candidate_manifest_status"] != "absent":
            print(f"  run={row['candidate_manifest_status']}")
        if row["db_error"]:
            print(f"  optuna=조회 실패 ({row['db_error']})")
            continue
        if not optuna:
            print("  optuna=-")
            continue

        counts = ", ".join(
            f"{key}:{value}"
            for key, value in sorted(optuna["counts"].items())
        ) or "-"
        print(f"  study={optuna['study_name']}")
        print(
            f"  trials={counts} | "
            f"best #{optuna['best_number']} "
            f"{_fmt_value(optuna['best_value'])}"
        )
        if optuna["latest_number"] is not None:
            progress = ""
            step = optuna["latest_intermediate_step"]
            if step is not None:
                current = step + 1
                total = row["progress_total"]
                progress = (
                    f" | {row['progress_unit']} "
                    f"{current}/{total if total else '?'}"
                )
                if optuna["latest_intermediate_value"] is not None:
                    progress += (
                        " | current "
                        f"{_fmt_value(optuna['latest_intermediate_value'])}"
                    )
            print(
                f"  latest=#{optuna['latest_number']} "
                f"{optuna['latest_state']}{progress}"
            )



def _clear_screen() -> None:
    # Git Bash/MinTTY may report isatty=False for Windows Python even though
    # ANSI cursor control is supported. Emit the clear sequence unconditionally.
    print("\033[2J\033[H", end="", flush=True)


def _progress_bar(
    current: int,
    total: int,
    width: int = 22,
    *,
    active: bool = False,
) -> str:
    if total <= 0:
        return "░" * width
    ratio = max(0.0, min(1.0, current / total))
    filled = int(width * ratio)
    head = 1 if active and filled < width and current > 0 else 0
    return (
        "█" * filled
        + ("▌" if head else "")
        + "░" * max(0, width - filled - head)
    )


def _compact_candidate_label(candidate_id: str) -> str:
    label = candidate_id
    replacements = (
        ("observed_weather_history", "weather+history"),
        ("history_calendar", "history+calendar"),
        ("_lookback_", " · lookback "),
    )
    for old, new in replacements:
        label = label.replace(old, new)
    return label


def _parse_time(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed


def _duration_text(
    started_at: str | None,
    completed_at: str | None = None,
) -> str:
    start = _parse_time(started_at)
    if start is None:
        return "-"
    end = _parse_time(completed_at) or datetime.now(timezone.utc)
    seconds = max(0, int((end - start).total_seconds()))
    hours, rem = divmod(seconds, 3600)
    minutes, secs = divmod(rem, 60)
    return (
        f"{hours:02d}:{minutes:02d}:{secs:02d}"
        if hours
        else f"{minutes:02d}:{secs:02d}"
    )


def _gpu_status() -> dict[str, str] | None:
    command = [
        "nvidia-smi",
        "--query-gpu=utilization.gpu,memory.used,memory.total,"
        "temperature.gpu,power.draw",
        "--format=csv,noheader,nounits",
    ]
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=1.5,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return None
    if result.returncode != 0 or not result.stdout.strip():
        return None
    first = result.stdout.strip().splitlines()[0]
    parts = [part.strip() for part in first.split(",")]
    if len(parts) != 5:
        return None
    return {
        "util": parts[0],
        "memory_used": parts[1],
        "memory_total": parts[2],
        "temperature": parts[3],
        "power": parts[4],
    }


def _box_line(text: str = "", width: int = 76) -> str:
    visible = text[:width]
    return f"│ {visible:<{width}} │"


def _section(title: str, width: int = 76) -> str:
    label = f" {title} "
    return "├" + "─" + label + "─" * max(0, width - len(label)) + "┤"


ANSI_RESET = "\033[0m"
ANSI_BOLD = "\033[1m"
ANSI_DIM = "\033[2m"
ANSI_CYAN = "\033[96m"
ANSI_GREEN = "\033[92m"
ANSI_YELLOW = "\033[93m"
ANSI_RED = "\033[91m"
ANSI_WHITE = "\033[97m"


def _paint(text: str, *codes: str) -> str:
    return "".join(codes) + text + ANSI_RESET


def _state_text(state: str) -> str:
    upper = state.upper()
    if upper == "RUNNING":
        return _paint("RUNNING", ANSI_YELLOW, ANSI_BOLD)
    if upper == "COMPLETE":
        return _paint("DONE", ANSI_GREEN, ANSI_BOLD)
    if upper == "PRUNED":
        return _paint("PRUNED", ANSI_DIM)
    if upper == "FAIL":
        return _paint("FAIL", ANSI_RED, ANSI_BOLD)
    return state


def _pct(current: int, total: int) -> int:
    return round(current * 100 / total) if total else 0


def _state_korean(state: str) -> str:
    upper = state.upper()
    return {
        "RUNNING": "학습 중",
        "COMPLETE": "완료",
        "PRUNED": "조기 종료",
        "FAIL": "실패",
        "WAITING": "대기",
    }.get(upper, state)


def _progress_label(unit: str) -> str:
    return {
        "epoch": "Epoch",
        "round": "Round",
        "step": "Step",
    }.get(unit, unit)


def _feature_description(candidate_id: str) -> tuple[str, str | None]:
    if candidate_id.startswith("observed_weather_history"):
        feature = "기상 + 발전이력"
    elif candidate_id.startswith("history_calendar"):
        feature = "발전이력 + 시간정보"
    else:
        feature = candidate_id

    lookback = None
    if "_lookback_" in candidate_id:
        raw = candidate_id.rsplit("_lookback_", 1)[1]
        lookback = raw[:-1] if raw.endswith("h") else raw
    return feature, lookback


def print_active_monitor(summary: dict[str, Any]) -> None:
    active_index = summary["active_candidate_index"]
    now = datetime.now().astimezone().strftime("%H:%M:%S")

    print(f"태양광 발전량 예측 학습                                      {now}")
    print("═" * 74)

    if active_index is None:
        print(f"현재 상태  {summary['overall_status']}")
        return

    row = next(
        item for item in summary["candidates"]
        if item["index"] == active_index
    )
    optuna = row["optuna"] or {}

    model_name = (
        "CNN-BiLSTM"
        if row["model"] == "cnn_bilstm"
        else "XGBoost"
        if row["model"] == "xgboost"
        else row["model"]
    )
    feature_name, lookback = _feature_description(row["candidate_id"])

    completed_total = sum(
        item["candidate_manifest_status"] == "completed"
        for item in summary["candidates"]
    )
    horizon_rows = [
        item
        for item in summary["candidates"]
        if item["horizon_hours"] == row["horizon_hours"]
    ]
    completed_horizon = sum(
        item["candidate_manifest_status"] == "completed"
        for item in horizon_rows
    )

    print(
        f"전체 진행  [{_progress_bar(completed_total, row['total'], 24)}] "
        f"{completed_total:>2}/{row['total']}  ({_pct(completed_total, row['total']):>3}%)"
    )
    print(
        f"{row['horizon_hours']}시간 구간  "
        f"[{_progress_bar(completed_horizon, len(horizon_rows), 24)}] "
        f"{completed_horizon}/{len(horizon_rows)}"
    )

    print("─" * 74)
    print(
        f"현재 후보  {row['index']}/{row['total']}   "
        f"{model_name}   ·   {row['horizon_hours']}시간 뒤 예측"
    )
    input_text = f"입력 조건  {feature_name}"
    if lookback is not None:
        input_text += f"   ·   과거 {lookback}시간 입력"
    print(input_text)

    latest_number = optuna.get("latest_number")
    trial_current = latest_number + 1 if latest_number is not None else 0
    counts = optuna.get("counts", {})
    complete_count = int(counts.get("COMPLETE", 0))
    pruned_count = int(counts.get("PRUNED", 0))
    failed_count = int(counts.get("FAIL", 0))

    step = optuna.get("latest_intermediate_step")
    step_current = step + 1 if step is not None else None
    step_total = row["progress_total"]
    progress_name = _progress_label(row["progress_unit"])

    if step_current is None and optuna.get("latest_state") == "RUNNING":
        progress_text = f"{progress_name}  첫 단계 학습 중"
    else:
        progress_text = (
            f"{progress_name}  {step_current if step_current is not None else '-'}"
            f"/{step_total or '?'}"
        )

    print("─" * 74)
    print(
        f"탐색 진행  {trial_current}/{row['max_trials']}   ·   {progress_text}"
    )
    print(
        f"탐색 집계  완료 {complete_count}   ·   "
        f"조기 종료 {pruned_count}   ·   실패 {failed_count}"
    )

    current_value = optuna.get("latest_intermediate_value")
    trial_best_value = optuna.get("latest_trial_best_value")
    trial_best_step = optuna.get("latest_trial_best_step")
    candidate_best_value = optuna.get("best_value")
    candidate_best_number = optuna.get("best_number")

    print("─" * 74)
    print("검증 MAE   낮을수록 좋음")
    print(
        "  현재 단계        "
        + (
            f"{current_value:.6f} MWh"
            if current_value is not None
            else "아직 검증 전"
        )
    )
    print(
        "  이번 탐색 최저   "
        + (
            f"{trial_best_value:.6f} MWh"
            f"   ({progress_name} {trial_best_step + 1})"
            if trial_best_value is not None and trial_best_step is not None
            else "아직 없음"
        )
    )
    print(
        "  이 후보 전체 최저 "
        + (
            f"{candidate_best_value:.6f} MWh"
            f"   (탐색 {candidate_best_number + 1})"
            if candidate_best_value is not None
            and candidate_best_number is not None
            else "아직 없음"
        )
    )

    recent = optuna.get("recent_trials", [])
    if recent:
        print("─" * 74)
        print("탐색 결과")
        print("  번호   상태        MAE            진행")
        for trial in sorted(recent, key=lambda item: int(item["number"])):
            number = int(trial["number"]) + 1
            state = _state_korean(str(trial["state"]))
            value = trial.get("value")
            value_text = (
                f"{value:.6f} MWh"
                if value is not None
                else "-"
            )
            last_step = trial.get("last_step")
            progress = (
                f"{progress_name} {int(last_step) + 1}/{step_total or '?'}"
                if last_step is not None
                else "-"
            )
            mark = (
                "  ★ 최저"
                if candidate_best_number is not None
                and int(trial["number"]) == candidate_best_number
                else ""
            )
            print(
                f"  {number:<5} {state:<10} {value_text:<14} {progress}{mark}"
            )



def _age_seconds(value: str | None) -> int | None:
    parsed = _parse_time(value)
    if parsed is None:
        return None
    return max(
        0,
        int((datetime.now(timezone.utc) - parsed).total_seconds()),
    )


def _age_text(seconds: int | None) -> str:
    if seconds is None:
        return "확인 불가"
    if seconds < 60:
        return f"{seconds}초 전"
    minutes, sec = divmod(seconds, 60)
    if minutes < 60:
        return f"{minutes}분 {sec}초 전"
    hours, minutes = divmod(minutes, 60)
    return f"{hours}시간 {minutes}분 전"


def _monitor_signature(summary: dict[str, Any]) -> tuple[Any, ...]:
    active = summary.get("active_candidate_index")
    if active is None:
        return (summary.get("overall_status"), None)
    row = next(
        item for item in summary["candidates"]
        if item["index"] == active
    )
    optuna = row.get("optuna") or {}
    return (
        active,
        optuna.get("latest_trial_id"),
        optuna.get("latest_number"),
        optuna.get("latest_state"),
        optuna.get("latest_intermediate_step"),
        optuna.get("latest_intermediate_value"),
        optuna.get("best_number"),
        optuna.get("best_value"),
        tuple(sorted((optuna.get("counts") or {}).items())),
    )


def _same_trial_line(summary: dict[str, Any]) -> str:
    active = summary.get("active_candidate_index")
    if active is None:
        return f"상태 확인 중 · {summary.get('overall_status')}"
    row = next(
        item for item in summary["candidates"]
        if item["index"] == active
    )
    optuna = row.get("optuna") or {}
    number = optuna.get("latest_number")
    trial = number + 1 if number is not None else "-"
    step = optuna.get("latest_intermediate_step")
    progress_name = _progress_label(row["progress_unit"])
    if step is None:
        progress = f"{progress_name} 첫 단계 계산 중"
    else:
        progress = f"{progress_name} {step + 1}/{row['progress_total'] or '?'}"
    heartbeat_age = _age_seconds(optuna.get("latest_heartbeat_at"))
    if heartbeat_age is not None and heartbeat_age > 180:
        health = f"⚠ heartbeat {_age_text(heartbeat_age)} · 멈춤 여부 확인 필요"
    elif heartbeat_age is not None:
        health = f"heartbeat {_age_text(heartbeat_age)} · 정상"
    else:
        health = "heartbeat 확인 불가"
    elapsed = _duration_text(optuna.get("latest_started_at"))
    return (
        f"탐색 {trial}/{row['max_trials']} 계속 계산 중 · {progress} · "
        f"재시작 없음 · 경과 {elapsed} · {health}"
    )


def _clear_keepalive_line(width: int = 140) -> None:
    print("\r" + " " * width + "\r", end="", flush=True)


def parse_args(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "실행 중 태양광 benchmark/Optuna 상태를 "
            "SQLite read-only로 조회합니다."
        )
    )
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="프로젝트 루트",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=Path("config/experiments/optimized.json"),
    )
    parser.add_argument(
        "--db",
        type=Path,
        default=Path("artifacts/optimization/solar_models.db"),
    )
    parser.add_argument(
        "--benchmarks",
        type=Path,
        default=Path("artifacts/benchmarks"),
    )
    parser.add_argument(
        "--lock",
        type=Path,
        default=Path("artifacts/.training.lock"),
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="기계 판독용 JSON 출력",
    )
    parser.add_argument(
        "--watch-seconds",
        type=float,
        default=None,
        help="N초마다 active 후보를 갱신합니다. 학습/DB는 변경하지 않습니다.",
    )
    return parser.parse_args(argv)


def main(argv: Iterable[str] | None = None) -> int:
    args = parse_args(argv)
    root = args.root.resolve()
    config = _resolve(root, args.config)
    db = _resolve(root, args.db)
    benchmarks = _resolve(root, args.benchmarks)
    lock = _resolve(root, args.lock)
    if args.watch_seconds is not None:
        if args.json:
            print("--json 과 --watch-seconds 는 함께 사용할 수 없습니다.", file=sys.stderr)
            return 2
        if args.watch_seconds <= 0:
            print("--watch-seconds 는 0보다 커야 합니다.", file=sys.stderr)
            return 2
        try:
            previous_signature = None
            while True:
                summary = summarize(
                    root,
                    config,
                    db,
                    benchmarks,
                    lock,
                )
                signature = _monitor_signature(summary)
                if signature != previous_signature:
                    _clear_keepalive_line()
                    _clear_screen()
                    print_active_monitor(summary)
                    previous_signature = signature
                else:
                    print(
                        "\r" + _same_trial_line(summary)[:138].ljust(138),
                        end="",
                        flush=True,
                    )
                time.sleep(args.watch_seconds)
        except KeyboardInterrupt:
            _clear_keepalive_line()
            return 0
        except (
            OSError,
            ValueError,
            KeyError,
            json.JSONDecodeError,
        ) as exc:
            print(
                f"상태 조회 실패: {type(exc).__name__}: {exc}",
                file=sys.stderr,
            )
            return 2

    try:
        summary = summarize(
            root,
            config,
            db,
            benchmarks,
            lock,
        )
    except (
        OSError,
        ValueError,
        KeyError,
        json.JSONDecodeError,
    ) as exc:
        print(
            f"상태 조회 실패: {type(exc).__name__}: {exc}",
            file=sys.stderr,
        )
        return 2

    if args.json:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    else:
        print_human(summary)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
