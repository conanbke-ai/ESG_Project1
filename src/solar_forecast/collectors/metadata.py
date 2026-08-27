from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Iterable

import pandas as pd

from .normalization import read_csv_with_fallback


def canonical_plant_name(value: object) -> str:
    """Return a comparison key while preserving site/unit distinctions."""

    text = str(value).lower()
    for token in ("태양광발전소", "태양광발전설비", "태양광", "발전소"):
        text = text.replace(token, "")
    return re.sub(r"[^0-9a-z가-힣#]", "", text)


def _unit_number(value: object) -> str | None:
    match = re.search(r"#\s*(\d+)", str(value))
    return match.group(1) if match else None


def _first_number(value: object) -> float | None:
    match = re.search(r"-?[0-9][0-9,]*(?:\.[0-9]+)?", str(value))
    return float(match.group(0).replace(",", "")) if match else None


def _capacity_mw(value: object, default_unit: str) -> float | None:
    text = str(value)
    number = _first_number(text)
    if number is None:
        return None
    unit_match = re.search(r"(?i)(mw|kw|w)\b", text)
    unit = unit_match.group(1).lower() if unit_match else default_unit.lower()
    # One KOSPO row says 997.56 W, while the same row states 340 W x 2,934
    # modules and 500 kW x 2 inverters. It is an obvious kW label typo.
    if unit == "w" and number < 10_000 and ("모듈" in text or "인버터" in text):
        unit = "kw"
    return number / {"w": 1_000_000, "kw": 1_000, "mw": 1}[unit]


@dataclass(frozen=True)
class PlantMetadata:
    company: str
    plant: str
    capacity_mw: float | None
    tilt_deg: float | None
    address: str | None
    unit: str | None = None

    @property
    def key(self) -> str:
        name = re.sub(r"#\d+$", "", canonical_plant_name(self.plant))
        return name


class PlantMetadataCatalog:
    """Normalize and match the four companies' public plant metadata tables."""

    aliases = {
        ("koen", "구미태양광"): "구미정수장",
        ("koen", "탑선태양광"): "탑선옥상형",
        ("kospo", "하동본부"): "하동화력",
        ("kospo", "부산본부"): "부산발전본부1400kw",
        ("kospo", "부산수처리장"): "부산수처리건물",
        ("kospo", "송당리"): "송당리제주",
        ("kospo", "신인천전망대"): "신인천법사면전망대",
        ("kospo", "신인천해수구취수구"): "신인천해수취수구",
        ("kospo", "하동보건소"): "하동군보건소",
        ("iwest", "(군산)삼랑진태양광"): "삼랑진 태양광 (FIT)",
        ("iwest", "영암에프원태양광b"): "영암F1 태양광",
        ("iwest", "안산연성정수장태양광"): "경기도 안산연성 태양광",
        ("iwest", "태안#9,10 수상태양광"): "태안수상태양광",
    }

    def __init__(self, records: Iterable[PlantMetadata]):
        self.records = tuple(records)

    @classmethod
    def from_directory(cls, directory: Path) -> "PlantMetadataCatalog":
        records: list[PlantMetadata] = []
        if not directory.exists():
            return cls(records)
        for path in sorted(directory.glob("*.csv")):
            frame = read_csv_with_fallback(path)
            columns = set(frame.columns)
            if {"발전소명", "설치용량", "설치각"}.issubset(columns):
                records.extend(cls._kospo(frame))
            elif {"사업명", "용량(kW)", "위치"}.issubset(columns):
                records.extend(cls._standard(frame, "koen", "용량(kW)", "kw", "위치"))
            elif {"사업명", "용량(kw)", "위치"}.issubset(columns):
                records.extend(cls._standard(frame, "ewp", "용량(kw)", "kw", "위치"))
            elif {"사업명", "용량(MW)", "소재지"}.issubset(columns):
                records.extend(cls._standard(frame, "iwest", "용량(MW)", "mw", "소재지"))
        return cls(records)

    @staticmethod
    def _kospo(frame: pd.DataFrame) -> list[PlantMetadata]:
        rows = []
        for row in frame.to_dict("records"):
            name = str(row["발전소명"]).strip()
            rows.append(
                PlantMetadata(
                    company="kospo",
                    plant=name,
                    unit=_unit_number(name),
                    capacity_mw=_capacity_mw(row.get("설치용량"), "kw"),
                    tilt_deg=_first_number(row.get("설치각")),
                    address=str(row.get("발전소 주소지")).strip() or None,
                )
            )
        return rows

    @staticmethod
    def _standard(
        frame: pd.DataFrame,
        company: str,
        capacity_column: str,
        capacity_unit: str,
        address_column: str,
    ) -> list[PlantMetadata]:
        rows = []
        for row in frame.to_dict("records"):
            name = str(row["사업명"]).strip()
            rows.append(
                PlantMetadata(
                    company=company,
                    plant=name,
                    unit=_unit_number(name),
                    capacity_mw=_capacity_mw(row.get(capacity_column), capacity_unit),
                    tilt_deg=None,
                    address=str(row.get(address_column)).strip() or None,
                )
            )
        return rows

    def lookup(
        self,
        company: str,
        plant: str,
        unit: str | None = None,
        *,
        aggregate: bool = False,
    ) -> PlantMetadata | None:
        query = self.aliases.get((company, plant), plant)
        query_key = re.sub(r"#\d+$", "", canonical_plant_name(query))
        candidates = [
            record
            for record in self.records
            if record.company == company
            and (
                record.key == query_key
                or (len(query_key) >= 4 and query_key in record.key)
                or (len(record.key) >= 4 and record.key in query_key)
            )
        ]
        if not candidates:
            return None
        if aggregate:
            capacities = [record.capacity_mw for record in candidates if record.capacity_mw is not None]
            tilts = [record.tilt_deg for record in candidates if record.tilt_deg is not None]
            addresses = [record.address for record in candidates if record.address]
            return PlantMetadata(
                company=company,
                plant=plant,
                capacity_mw=sum(capacities) if capacities else None,
                tilt_deg=sum(tilts) / len(tilts) if tilts else None,
                address=addresses[0] if addresses else None,
            )
        unit_text = re.sub(r"\.0$", "", str(unit).strip()) if unit is not None else None
        exact_unit = [record for record in candidates if record.unit == unit_text]
        if len(exact_unit) == 1:
            return exact_unit[0]
        no_unit = [record for record in candidates if record.unit is None]
        return no_unit[0] if len(candidates) == 1 and len(no_unit) == 1 else None

    def enrich(self, frame: pd.DataFrame, *, aggregate: bool = False) -> pd.DataFrame:
        result = frame.copy()
        keys = result[["company", "plant", "unit"]].drop_duplicates()
        metadata_rows: list[dict[str, object]] = []
        for row in keys.itertuples(index=False):
            company, plant, unit = str(row.company), str(row.plant), str(row.unit)
            record = self.lookup(company, plant, unit, aggregate=aggregate)
            if record:
                metadata_rows.append(
                    {
                        "company": company,
                        "plant": plant,
                        "unit": unit,
                        "_capacity_mw": record.capacity_mw,
                        "_tilt_deg": record.tilt_deg,
                        "_address": record.address,
                    }
                )
        if not metadata_rows:
            return result
        metadata = pd.DataFrame(metadata_rows)
        result = result.merge(metadata, on=["company", "plant", "unit"], how="left", validate="many_to_one")
        for target, fallback in (
            ("capacity_mw", "_capacity_mw"),
            ("tilt_deg", "_tilt_deg"),
            ("address", "_address"),
        ):
            result[target] = result[target].where(result[target].notna(), result[fallback])
        return result.drop(columns=["_capacity_mw", "_tilt_deg", "_address"])
