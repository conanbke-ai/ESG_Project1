from __future__ import annotations

import codecs
from dataclasses import asdict, dataclass
import hashlib
from itertools import chain
from pathlib import Path

import pandas as pd


@dataclass(frozen=True)
class CsvArtifactAudit:
    path: str
    encoding: str
    byte_order_mark: bool
    bytes: int
    sha256: str

    def as_dict(self) -> dict[str, object]:
        return asdict(self)


def inspect_csv_artifact(path: Path) -> CsvArtifactAudit:
    """Hash and identify a Korean CSV in one bounded-memory sequential scan."""

    digest = hashlib.sha256()
    decoders = {
        "utf-8": codecs.getincrementaldecoder("utf-8")(errors="strict"),
        "cp949": codecs.getincrementaldecoder("cp949")(errors="strict"),
    }
    valid = {encoding: True for encoding in decoders}
    with path.open("rb") as source:
        prefix = source.read(3)
        for chunk in chain((prefix,), iter(lambda: source.read(1024 * 1024), b"")):
            digest.update(chunk)
            for encoding, decoder in decoders.items():
                if not valid[encoding]:
                    continue
                try:
                    decoder.decode(chunk)
                except UnicodeDecodeError:
                    valid[encoding] = False
    for encoding, decoder in decoders.items():
        if valid[encoding]:
            try:
                decoder.decode(b"", final=True)
            except UnicodeDecodeError:
                valid[encoding] = False
    has_bom = prefix == codecs.BOM_UTF8
    if has_bom:
        encoding = "utf-8-sig"
    elif valid["utf-8"]:
        encoding = "utf-8"
    elif valid["cp949"]:
        encoding = "cp949"
    else:
        encoding = "unknown"
    return CsvArtifactAudit(
        path=str(path),
        encoding=encoding,
        byte_order_mark=has_bom,
        bytes=path.stat().st_size,
        sha256=digest.hexdigest(),
    )


def write_standardized_csv(frame: pd.DataFrame, destination: Path) -> Path:
    """Atomically write a Silver CSV using Excel-compatible UTF-8 with BOM."""

    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_suffix(destination.suffix + ".part")
    frame.to_csv(temporary, index=False, encoding="utf-8-sig")
    temporary.replace(destination)
    return destination
