# src/core/cleanup.py
from __future__ import annotations

from pathlib import Path
from typing import Iterable, Set, Tuple


def cleanup_output_dir_keep_pdf_xml(
    output_dir: Path,
    keep_exts: Iterable[str] = (".pdf", ".xml"),
    keep_paths: Iterable[Path] = (),
) -> Tuple[int, int]:
    """
    Elimina recursivamente archivos dentro de output_dir que NO estén en keep_exts
    y que NO estén en keep_paths.

    Retorna: (deleted_files, deleted_dirs)
    """
    output_dir = Path(output_dir).resolve()
    keep: Set[Path] = {Path(p).resolve() for p in keep_paths}
    keep_exts = {e.lower() if e.startswith(".") else f".{e.lower()}" for e in keep_exts}

    deleted_files = 0
    deleted_dirs = 0

    if not output_dir.exists():
        return (0, 0)

    for p in output_dir.rglob("*"):
        if not p.is_file():
            continue

        rp = p.resolve()
        if rp in keep:
            continue
        if p.suffix.lower() in keep_exts:
            continue

        try:
            p.unlink()
            deleted_files += 1
        except Exception:
            pass

    for d in sorted([x for x in output_dir.rglob("*") if x.is_dir()], reverse=True):
        try:
            if not any(d.iterdir()):
                d.rmdir()
                deleted_dirs += 1
        except Exception:
            pass

    return deleted_files, deleted_dirs
