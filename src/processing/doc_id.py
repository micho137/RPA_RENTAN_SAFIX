from __future__ import annotations
import os
import re
from typing import Optional, Tuple
from dotenv import load_dotenv

load_dotenv()


def _load_valid_series() -> tuple[str, ...]:
    raw = os.getenv("DOC_ID_VALID_SERIES", "DEFL,FPFL")
    items = [x.strip().upper() for x in str(raw).split(",")]
    items = [x for x in items if x]
    if not items:
        items = ["DEFL", "FPFL"]
    seen = set()
    ordered: list[str] = []
    for s in items:
        if s not in seen:
            seen.add(s)
            ordered.append(s)
    return tuple(ordered)


VALID_SERIES = set(_load_valid_series())
_SERIES_PATTERN = "|".join(re.escape(s) for s in sorted(VALID_SERIES, key=len, reverse=True))
DOC_ID_RE = re.compile(rf"\b(?P<serie>{_SERIES_PATTERN})-?(?P<num>\d{{8}})\b", re.IGNORECASE)

def parse_any_id(raw: str) -> Optional[Tuple[str, str, str]]:
    """
    Dada una cadena, extrae (serie, numero, document_id_normalizado).
    document_id_normalizado siempre con guion: 'SERIE-########'.
    """
    if not raw:
        return None
    m = DOC_ID_RE.search(raw)
    if not m:
        return None
    serie = m.group("serie").upper()
    numero = m.group("num")
    if serie not in VALID_SERIES:
        return None
    norm = f"{serie}-{numero}"
    return serie, numero, norm

def normalize_id(raw: str) -> Optional[str]:
    """Devuelve 'SERIE-########' o None si no coincide."""
    out = parse_any_id(raw)
    return out[2] if out else None

def split_parts(document_id: str) -> Tuple[Optional[str], Optional[str]]:
    """De 'SERIE-########' devuelve ('SERIE','########')."""
    m = DOC_ID_RE.search(document_id or "")
    if not m:
        return None, None
    return m.group("serie").upper(), m.group("num")
