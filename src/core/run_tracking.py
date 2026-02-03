# src/core/run_tracking.py
from __future__ import annotations

from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any

from openpyxl import Workbook
from openpyxl import load_workbook
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


@dataclass
class MissingPlateRow:
    timestamp: str
    document_id: str
    placa: str
    total: Optional[int]
    key: str
    status: str  # MISSING
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
        self.missing_plates: List[MissingPlateRow] = []

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
    ) -> ProcessRow:
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
        self.processed.append(row)
        return row

    def add_missing_plate(
        self,
        *,
        document_id: str,
        placa: str,
        total: Optional[int],
        key: str,
        error: str = "",
    ) -> MissingPlateRow:
        row = MissingPlateRow(
            timestamp=_now_iso(),
            document_id=document_id,
            placa=placa,
            total=total,
            key=key,
            status="MISSING",
            error=error,
        )
        self.missing_plates.append(row)
        return row

    def save(self) -> dict[str, Path]:
        out = {}

        # p1 = self.logs_dir / "descargados.xlsx"
        # self._write_xlsx(p1, "descargados", [asdict(x) for x in self.downloads])
        # out["descargados"] = p1

        p2 = self.logs_dir / "procesadas.xlsx"
        self._write_xlsx(p2, "procesadas", [asdict(x) for x in self.processed])
        out["procesadas"] = p2

        p3 = self.logs_dir / "placas_no_encontradas.xlsx"
        self._write_xlsx(p3, "placas_no_encontradas", [asdict(x) for x in self.missing_plates])
        out["placas_no_encontradas"] = p3

        return out

    def append_processed_row(self, row: ProcessRow) -> None:
        """
        Agrega una fila al Excel de procesadas en tiempo real.
        Si el archivo no existe, crea cabeceras.
        """
        path = self.logs_dir / "procesadas.xlsx"
        headers = list(asdict(row).keys())

        if path.exists():
            wb = load_workbook(filename=str(path))
            ws = wb.active
            # Si no hay cabeceras, agregarlas
            if ws.max_row == 0 or (ws.max_row == 1 and ws.cell(1, 1).value is None):
                ws.append(headers)
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "procesadas"
            ws.append(headers)

        ws.append([getattr(row, h, "") for h in headers])
        wb.save(path)

    def append_missing_plate_row(self, row: MissingPlateRow) -> None:
        """
        Agrega una fila al Excel de placas no encontradas en tiempo real.
        """
        path = self.logs_dir / "placas_no_encontradas.xlsx"
        headers = list(asdict(row).keys())

        if path.exists():
            wb = load_workbook(filename=str(path))
            ws = wb.active
            if ws.max_row == 0 or (ws.max_row == 1 and ws.cell(1, 1).value is None):
                ws.append(headers)
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "placas_no_encontradas"
            ws.append(headers)

        ws.append([getattr(row, h, "") for h in headers])
        wb.save(path)

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
