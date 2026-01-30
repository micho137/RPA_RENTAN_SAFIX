from __future__ import annotations
import re
from typing import Optional, Tuple

#VALID_SERIES = {"DEFL", "DENC"}
VALID_SERIES = {"DEFL"}

# Match "DEFL" o "DENC" con o sin guion y 8 dígitos
DOC_ID_RE = re.compile(r"\b(?P<serie>DEFL|DENC)-?(?P<num>\d{8})\b", re.IGNORECASE)

def parse_any_id(raw: str) -> Optional[Tuple[str, str, str]]:
    """
    Dada una cadena, extrae (serie, numero, document_id_normalizado).
    document_id_normalizado siempre con guion: 'DEFL-########' o 'DENC-########'.
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
