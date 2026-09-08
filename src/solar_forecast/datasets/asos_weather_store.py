"""Merge API observations into the annual ASOS CSV contract consumed by Gold."""
from __future__ import annotations

from contextlib import contextmanager
import csv
from datetime import datetime
import os
from pathlib import Path

from solar_forecast.infrastructure.artifact_store import replace_file_atomic

ASOS_API_COLUMNS = {
    "stnId": "지점", "stnNm": "지점명", "tm": "일시",
    "ta": "기온(°C)", "rn": "강수량(mm)", "ws": "풍속(m/s)",
    "hm": "습도(%)", "ss": "일조(hr)", "icsr": "일사(MJ/m2)",
    "dc10Tca": "전운량(10분위)", "dc10LmcsCa": "중하층운량(10분위)",
}


@contextmanager
def exclusive_asos_collection(weather_root: Path):
    """Serialize API collection and annual-file publication across local workers."""
    weather_root.mkdir(parents=True, exist_ok=True)
    lock = weather_root / ".asos_collection.lock"
    try:
        descriptor = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
    except FileExistsError:
        raise RuntimeError("Another ASOS collection holds the weather lock; inspect the lock owner before recovery") from None
    try:
        with os.fdopen(descriptor, "w") as stream:
            stream.write(str(os.getpid()))
        yield
    finally:
        lock.unlink(missing_ok=True)


def observation_to_weather_row(item: dict) -> dict[str, str]:
    """Preserve QC flags and missing values while translating provider field names."""
    row = {column: str(item.get(field) if item.get(field) is not None else "")
           for field, column in ASOS_API_COLUMNS.items() if field in item}
    row["일시"] = datetime.fromisoformat(row["일시"]).strftime("%Y-%m-%d %H:%M")
    for field, column in ASOS_API_COLUMNS.items():
        flag_name = field + "Qcflg"
        if flag_name in item:
            flag = str(item[flag_name] if item[flag_name] is not None else "")
            row[flag_name] = flag
            if flag not in {"", "0"}:
                row[column] = ""
    return row


def _read_annual_weather(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    for encoding in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            with path.open(encoding=encoding, newline="") as stream:
                reader = csv.DictReader(stream)
                columns = [str(column).strip() for column in reader.fieldnames or []]
                rows = [{str(key).strip(): value for key, value in row.items()} for row in reader]
            if not {"지점", "일시"}.issubset(columns):
                raise ValueError(f"Existing ASOS annual CSV lacks station/time keys: {path}")
            return columns, rows
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Cannot decode ASOS annual CSV: {path}")


def _weather_row_key(row: dict[str, str]) -> tuple[str, str]:
    return str(int(row["지점"])), datetime.fromisoformat(row["일시"]).strftime("%Y-%m-%d %H:%M")


def merge_asos_observations(weather_root: Path, items: list[dict]) -> list[Path]:
    """Upsert station/hour keys, retaining other stations and unknown CSV columns.

    The caller holds exclusive_asos_collection. Annual CSVs remain CP949 for the
    existing browser collector; Bronze responses retain the full provider JSON.
    """
    return merge_weather_rows(weather_root, [observation_to_weather_row(item) for item in items])


def merge_weather_rows(weather_root: Path, rows: list[dict[str, str]]) -> list[Path]:
    """Publish rows from either API or browser using the same keyed column merge."""
    by_year: dict[int, list[dict[str, str]]] = {}
    for row in rows:
        by_year.setdefault(int(row["일시"][:4]), []).append(row)
    paths = []
    for year, incoming in sorted(by_year.items()):
        path = weather_root / f"OBS_ASOS_TIM_{year}.csv"
        columns, existing = _read_annual_weather(path) if path.exists() else (list(ASOS_API_COLUMNS.values()), [])
        columns = list(dict.fromkeys([*columns, *ASOS_API_COLUMNS.values(), *(key for row in incoming for key in row)]))
        merged = {_weather_row_key(row): row for row in existing}
        for row in incoming:
            key = _weather_row_key(row)
            merged[key] = {**merged.get(key, {}), **row}
        temporary = path.with_name(path.name + ".part")
        with temporary.open("w", encoding="cp949", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=columns)
            writer.writeheader()
            writer.writerows(merged[key] for key in sorted(merged, key=lambda key: (key[1], int(key[0]))))
        replace_file_atomic(temporary, path)
        paths.append(path)
    return paths
