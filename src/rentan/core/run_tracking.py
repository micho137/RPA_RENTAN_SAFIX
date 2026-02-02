from __future__ import annotations

from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any

from openpyxl import Workbook, load_workbook


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


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
    Log incremental:
      - output/logs/procesadas.xlsx
    """

    def __init__(self, output_dir: Path) -> None:
        self.output_dir = Path(output_dir)
        self.logs_dir = self.output_dir / "logs"
        self.logs_dir.mkdir(parents=True, exist_ok=True)

        self.path_procesadas = self.logs_dir / "procesadas.xlsx"
        self._ensure_workbook()

    def _ensure_workbook(self) -> None:
        if self.path_procesadas.exists():
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "procesadas"

        headers = list(asdict(ProcessRow(
            timestamp_start="",
            timestamp_end="",
            document_id="",
            placa="",
            total=None,
            key="",
            status="",
            error="",
        )).keys())

        ws.append(headers)
        wb.save(self.path_procesadas)

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
        row = ProcessRow(
            timestamp_start=timestamp_start,
            timestamp_end=timestamp_end,
            document_id=document_id,
            placa=placa,
            total=total,
            key=key,
            status=status,
            error=error,
        )
        self._append_row(asdict(row))

    def _append_row(self, row: Dict[str, Any]) -> None:
        wb = load_workbook(self.path_procesadas)
        ws = wb.active  # "procesadas"

        # Asume headers en fila 1
        headers = [c.value for c in ws[1]]
        ws.append([row.get(h, "") for h in headers])

        wb.save(self.path_procesadas)
