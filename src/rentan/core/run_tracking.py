from __future__ import annotations

from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any

from openpyxl import Workbook, load_workbook
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
class ProcessRow:
    timestamp_start: str
    timestamp_end: str
    document_id: str
    placa: str
    total: Optional[int]
    key: str
    status: str
    error: str = ""


class RunTracker:
    """
    Genera y mantiene actualizado:
      output/logs/procesadas.xlsx

    Escribe INCREMENTALMENTE por cada factura.
    """

    def __init__(self, output_dir: Path) -> None:
        self.output_dir = Path(output_dir)
        self.logs_dir = self.output_dir / "logs"
        self.logs_dir.mkdir(parents=True, exist_ok=True)

        self.path_procesadas = self.logs_dir / "procesadas.xlsx"

        if not self.path_procesadas.exists():
            self._init_workbook()

    def _init_workbook(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "procesadas"

        headers = list(ProcessRow.__dataclass_fields__.keys())
        ws.append(headers)

        _autosize(ws)
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

    def _append_row(self, data: Dict[str, Any]) -> None:
        if self.path_procesadas.exists():
            wb = load_workbook(self.path_procesadas)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "procesadas"
            ws.append(list(data.keys()))

        ws.append([data.get(h, "") for h in data.keys()])
        _autosize(ws)
        wb.save(self.path_procesadas)
