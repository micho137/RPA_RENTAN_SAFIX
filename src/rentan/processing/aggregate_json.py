from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional, Iterable


@dataclass
class AggregateResult:
    total_ids: int
    written_by_id: Path


def _safe_read_json(path: Path) -> Optional[Dict[str, Any]]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def build_invoices_by_id(
    json_root: Path,
    out_by_id_path: Path,
    include_source_path: bool = True,
    exclude_names: Iterable[str] = ("all_invoices_by_id.json", "all_invoices.json"),
) -> AggregateResult:
    """
    Recorre todos los *.json en json_root (recursivo) y genera SOLO:
      - Un JSON dict indexado por document_id: out_by_id_path

    Evita auto-incluir agregados previos (exclude_names).
    Si el mismo document_id aparece varias veces, el último gana.
    """
    json_root = Path(json_root)
    out_by_id_path = Path(out_by_id_path)

    by_id: Dict[str, Dict[str, Any]] = {}
    exclude = set(exclude_names)

    for p in sorted(json_root.rglob("*.json")):
        if not p.is_file():
            continue
        if p.name in exclude:
            continue

        data = _safe_read_json(p)
        if not isinstance(data, dict):
            continue

        doc_id = ((data.get("document") or {}).get("document_id")) or ""
        doc_id = str(doc_id).strip()
        if not doc_id:
            continue

        if include_source_path:
            data["_source_json"] = str(p)

        by_id[doc_id] = data

    out_by_id_path.parent.mkdir(parents=True, exist_ok=True)
    out_by_id_path.write_text(
        json.dumps(by_id, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    return AggregateResult(total_ids=len(by_id), written_by_id=out_by_id_path)
