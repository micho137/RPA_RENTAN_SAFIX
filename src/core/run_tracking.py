# src/core/run_tracking.py
from __future__ import annotations

from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any

from openpyxl import Workbook
from openpyxl.utils import get_column_letter


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _autosize(ws):
    for col in range(1, ws.max_column + 1):
        max_len = 0
        col_letter = get_column_letter(col)
        for cell in ws[col_letter]:
            v = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(v))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 70)


@dataclass
class DownloadRow:
    timestamp: str
    original_name: str
    saved_path: str
    size_bytes: Optional[int] = None
    status: str = "OK"
    error: str = ""


@dataclass
class ProcessRow:
    timestamp_start: str
    timestamp_end: str
    document_id: str
    placa: str
    total: Optional[int]
    key: str
    status: str  # OK / FAIL
    error: str = ""


class RunTracker:
    """
    Logs:
      - output/logs/descargados.xlsx
      - output/logs/procesadas.xlsx
    """

    def __init__(self, output_dir: Path) -> None:
        self.output_dir = Path(output_dir)
        self.logs_dir = self.output_dir / "logs"
        self.logs_dir.mkdir(parents=True, exist_ok=True)

        self.downloads: List[DownloadRow] = []
        self.processed: List[ProcessRow] = []

    def add_download(
        self,
        *,
        original_name: str,
        saved_path: Path,
        size_bytes: Optional[int] = None,
        status: str = "OK",
        error: str = "",
    ) -> None:
        self.downloads.append(
            DownloadRow(
                timestamp=_now_iso(),
                original_name=original_name,
                saved_path=str(saved_path),
                size_bytes=size_bytes,
                status=status,
                error=error,
            )
        )

    def add_processed(
        self,
        *,
        timestamp_start: str,
        timestamp_end: str,
        document_id: str,
        placa: str,
        total: Optional[int],
        key: str,
        status: str,
        error: str = "",
    ) -> None:
        self.processed.append(
            ProcessRow(
                timestamp_start=timestamp_start,
                timestamp_end=timestamp_end,
                document_id=document_id,
                placa=placa,
                total=total,
                key=key,
                status=status,
                error=error,
            )
        )

    def save(self) -> dict[str, Path]:
        out = {}

        # p1 = self.logs_dir / "descargados.xlsx"
        # self._write_xlsx(p1, "descargados", [asdict(x) for x in self.downloads])
        # out["descargados"] = p1

        p2 = self.logs_dir / "procesadas.xlsx"
        self._write_xlsx(p2, "procesadas", [asdict(x) for x in self.processed])
        out["procesadas"] = p2

        return out

    def _write_xlsx(self, path: Path, sheet_name: str, rows: List[Dict[str, Any]]) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name

        if not rows:
            ws.append(["sin datos"])
            wb.save(path)
            return

        headers = list(rows[0].keys())
        ws.append(headers)
        for r in rows:
            ws.append([r.get(h, "") for h in headers])

        _autosize(ws)
        wb.save(path)
