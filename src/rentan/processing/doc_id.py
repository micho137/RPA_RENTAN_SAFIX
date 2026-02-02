from __future__ import annotations
import re
from typing import Optional, Tuple

DOC_ID_RE = re.compile(r"\b(?P<serie>DEFL)-?(?P<num>\d{8})\b", re.IGNORECASE)

def parse_any_id(raw: str) -> Optional[Tuple[str, str, str]]:
    if not raw:
        return None
    m = DOC_ID_RE.search(raw)
    if not m:
        return None
    serie = "DEFL"
    numero = m.group("num")
    norm = f"{serie}-{numero}"
    return serie, numero, norm

def normalize_id(raw: str) -> Optional[str]:
    out = parse_any_id(raw)
    return out[2] if out else None

def split_parts(document_id: str) -> Tuple[Optional[str], Optional[str]]:
    m = DOC_ID_RE.search(document_id or "")
    if not m:
        return None, None
    return "DEFL", m.group("num")
