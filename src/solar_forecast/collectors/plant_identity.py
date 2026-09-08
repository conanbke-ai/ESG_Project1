"""Evidence-backed physical identities shared by new and archived Silver reads."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd


KOSPO_IDENTITY_RULE = "kospo.busan-rail-2.v1"
KOSPO_IDENTITY_EVIDENCE = "https://www.data.go.kr/data/15156688/fileData.do"
KOSPO_METADATA_EVIDENCE = (
    "file/solar_data_file/location/"
    "한국남부발전(주)_태양광발전기 사양정보_20250630.csv"
)


def resolve_generation_identity(frame: pd.DataFrame) -> pd.DataFrame:
    """Resolve only the reviewed department/generator pair, without guessing.

    Dataset 15156688 calls the department ``발전소명`` and the physical asset
    ``발전기명``. The specification table names that asset ``부산철도 2``.
    A department-wide alias would merge different generators. Keep the meter
    and original generator label in plant_id/unit, and leave unknown pairs alone.
    No capacity, coordinates, weather mapping or generation values are inferred.
    """

    if not {"company", "plant", "unit"}.issubset(frame.columns):
        return frame
    generator = frame["unit"].astype("string").str.replace(r"\s+", "", regex=True)
    matches = (
        frame["company"].eq("kospo")
        & frame["plant"].eq("신재생사업본부")
        & generator.eq("부산철도태양광#2")
    ).fillna(False)
    if "energy_source" in frame:
        matches &= frame["energy_source"].eq("solar").fillna(False)
    if not matches.any():
        return frame

    result = frame.copy()
    result.loc[matches, "plant"] = "부산철도 2"
    if "plant_id" in result:
        # Validate the original generator too, then keep its meter as a suffix.
        ids = result.loc[matches, "plant_id"].astype("string")
        parts = ids.str.rsplit("#", n=1)
        expected = "kospo:신재생사업본부#" + frame.loc[matches, "unit"].astype("string")
        valid = parts.str[0].eq(expected) & parts.str[1].str.strip().ne("")
        if not valid.fillna(False).all():
            raise ValueError("KOSPO plant_id conflicts with the reviewed source identity")
        result.loc[matches, "plant_id"] = (
            "kospo:부산철도 2#부산 철도태양광 #2#" + parts.str[1]
        )
        if "timestamp" in result:
            collisions = result.duplicated(["timestamp", "plant_id"], keep=False)
            if (collisions & matches).any():
                raise ValueError("KOSPO identity resolution produces duplicate hourly keys")
    result.attrs["plant_identity_resolutions"] = [
        {
            "rule": KOSPO_IDENTITY_RULE,
            "company": "kospo",
            "source_plant": "신재생사업본부",
            "source_unit": "부산 철도태양광 #2",
            "canonical_plant": "부산철도 2",
            "rows": int(matches.sum()),
            "evidence_url": KOSPO_IDENTITY_EVIDENCE,
            "metadata_source": KOSPO_METADATA_EVIDENCE,
        }
    ]
    return result


def read_generation_partition(path: Path, columns: Iterable[str]) -> pd.DataFrame:
    """Read the optional generator label before projecting registry/Gold columns.

    Historical partitions without ``unit`` retain their existing behavior. This
    read-only projection also repairs cached collector Silver without rewriting
    its bytes or requiring a fresh official download.
    """

    selected = list(columns)
    accepted_columns = {*selected, "unit"}
    frame = pd.read_csv(path, usecols=lambda column: column in accepted_columns)
    missing = set(selected) - set(frame.columns)
    if missing:
        raise ValueError(f"Generation partition columns are missing: {sorted(missing)}")
    return resolve_generation_identity(frame)[selected]
