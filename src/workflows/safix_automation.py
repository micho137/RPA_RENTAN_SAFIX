from __future__ import annotations

import json
import os
import re
import time
from datetime import datetime
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import pyautogui
from openpyxl import load_workbook
from pywinauto import Application, Desktop

from src.config import settings
from src.core.com_init import com_initialized
from src.outlook.client import OutlookClient
from src.ui.overlay_status import StatusOverlay

# Mensaje simple cuando no hay facturas
try:
    import tkinter as _tk
    from tkinter import messagebox as _messagebox
except Exception:  # pragma: no cover
    _tk = None
    _messagebox = None



# =========================
# Regex / Utilidades
# =========================
PLACA_REGEX = re.compile(r"\bPLACA\b\s*[:\-]?\s*([A-Z0-9]{5,8})\b", re.IGNORECASE)


def _strip_quotes(s: str) -> str:
    """Quita comillas simples/dobles envolventes si existen."""
    if not s:
        return s
    s = s.strip()
    if (s.startswith("'") and s.endswith("'")) or (s.startswith('"') and s.endswith('"')):
        return s[1:-1].strip()
    return s


def normalize_placa(s: str) -> str:
    """Normaliza placa a mayúsculas y sin espacios."""
    return re.sub(r"\s+", "", str(s or "").strip().upper())


# =========================
# Loaders
# =========================
def load_invoices_by_id(path: Path) -> Dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"No existe el agregado invoices_by_id: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("El archivo all_invoices_by_id.json no es un dict.")
    return data


def load_invoices_from_json_dir(json_root: Path) -> Dict[str, Any]:
    """
    Carga todos los JSON individuales en un dict {key: data},
    donde key es el nombre del archivo (stem).
    """
    json_root = Path(json_root)
    if not json_root.exists():
        raise FileNotFoundError(f"No existe el directorio JSON: {json_root}")

    out: Dict[str, Any] = {}
    for p in sorted(json_root.rglob("*.json")):
        if not p.is_file():
            continue
        if p.name in {"all_invoices_by_id.json", "all_invoices.json"}:
            continue
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        if isinstance(data, dict):
            out[p.stem] = data
    return out


def load_attachment_manifest(download_dir: Path) -> Dict[str, str]:
    path = Path(download_dir) / "_manifest.json"
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _zip_stem_from_source_path(source_path: str) -> Optional[str]:
    try:
        p = Path(source_path)
        extract_root = settings.extract_dir
        rel = p.relative_to(extract_root)
        return rel.parts[0] if rel.parts else None
    except Exception:
        # fallback: buscar carpeta "extract" en el path
        parts = Path(source_path).parts
        if "extract" in parts:
            idx = parts.index("extract")
            if idx + 1 < len(parts):
                return parts[idx + 1]
        return None


def load_processed_ids_from_log(xlsx_path: Path) -> set[str]:
    """
    Lee procesadas.xlsx y retorna set de document_id con status OK.
    """
    if not xlsx_path.exists():
        return set()

    try:
        wb = load_workbook(filename=str(xlsx_path), read_only=True, data_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if not header:
            return set()

        header_map = {str(h).strip().lower(): idx for idx, h in enumerate(header) if h is not None}
        doc_idx = header_map.get("document_id")
        status_idx = header_map.get("status")

        if doc_idx is None:
            return set()

        processed = set()
        for row in rows:
            if not row:
                continue
            doc = row[doc_idx] if doc_idx < len(row) else None
            status = row[status_idx] if status_idx is not None and status_idx < len(row) else None
            if doc and (status is None or str(status).strip().upper() == "OK"):
                processed.add(str(doc).strip())
        return processed
    except Exception:
        return set()


def _show_no_invoices_message():
    if _tk is None or _messagebox is None:
        return
    try:
        root = _tk.Tk()
        root.withdraw()
        _messagebox.showinfo("RENTAN", "No hay facturas para procesar.")
        root.destroy()
    except Exception:
        pass


def load_plate_catalog_from_excel(excel_path: Path) -> Dict[str, Dict[str, str]]:
    """
    Lee el Excel y retorna un catálogo por placa:
      {
        "ABC123": {"interface": "XXXX", "centro_costos": "YYYY"},
        ...
      }
    Requiere columnas: PLACA, CENTRO DE COSTOS, INTERFACE
    """
    if not excel_path.exists():
        raise FileNotFoundError(f"No existe el Excel de placas: {excel_path}")

    wb = load_workbook(filename=str(excel_path), data_only=True)
    ws = wb.active

    headers: Dict[str, int] = {}
    for col_idx, cell in enumerate(ws[1], start=1):
        h = str(cell.value or "").strip().upper()
        if h:
            headers[h] = col_idx

    required = ["PLACA", "CENTRO DE COSTOS", "INTERFACE", "DOBLE CC"]
    missing = [r for r in required if r not in headers]
    if missing:
        raise ValueError(
            f"El Excel no tiene columnas requeridas {missing}. "
            f"Encontradas: {list(headers.keys())}"
        )

    catalog: Dict[str, Dict[str, str]] = {}
    for row_idx in range(2, ws.max_row + 1):
        placa_val = ws.cell(row=row_idx, column=headers["PLACA"]).value
        if not placa_val:
            continue

        placa = normalize_placa(placa_val)
        if not placa:
            continue

        centro_costos = ws.cell(row=row_idx, column=headers["CENTRO DE COSTOS"]).value
        interface = ws.cell(row=row_idx, column=headers["INTERFACE"]).value
        doble_cc = ws.cell(row=row_idx, column=headers["DOBLE CC"]).value

        catalog[placa] = {
            "centro_costos": str(centro_costos or "").strip(),
            "interface": str(interface or "").strip(),
            "doble_cc": str(doble_cc or "").strip(),
        }

    return catalog


# =========================
# Extracción de campos (doc_id, placa, total)
# =========================
def extract_doc_plate_and_total(
    invoice: Dict[str, Any],
    fallback_plate: str,
    key_fallback_doc_id: Optional[str] = None,
) -> Tuple[str, str, int]:
    # -------- document_id --------
    doc_id = ((invoice.get("document") or {}).get("document_id")) or ""
    doc_id = str(doc_id).strip()

    if not doc_id and key_fallback_doc_id:
        doc_id = str(key_fallback_doc_id).strip()

    if not doc_id:
        raise ValueError("Factura sin document_id (ni en JSON ni en la key del agregado).")

    doc_id = doc_id.replace("-", "").strip()

    # -------- placa --------
    placa = invoice.get("placa") or (invoice.get("vehiculo") or {}).get("placa") or ""
    placa = str(placa).strip()

    if not placa:
        for item in (invoice.get("detalle") or []):
            desc = str((item or {}).get("descripcion") or "")
            m = PLACA_REGEX.search(desc)
            if m:
                placa = m.group(1).strip()
                break

    if not placa:
        m2 = PLACA_REGEX.search(str(fallback_plate))
        placa = (m2.group(1).strip() if m2 else str(fallback_plate).strip())

    placa = normalize_placa(placa)

    # -------- total (peaje) --------
    total_val = ((invoice.get("totales") or {}).get("total"))
    if total_val is None:
        raise ValueError(f"Factura {doc_id}: no trae totales.total")

    try:
        total_int = int(float(str(total_val).strip()))
    except Exception:
        raise ValueError(f"Factura {doc_id}: totales.total inválido: {total_val!r}")

    return doc_id, placa, total_int


# =========================
# Config SAFIX
# =========================
@dataclass(frozen=True)
class SafixConfig:
    # App/ventana
    jnlp_path: Path
    main_window_title_re: str

    # UI (icons)
    tesoreria_icon: Path
    valores_icon: Path
    z_icon: Path
    engranes_icon: Path

    # Credenciales / negocio
    user: str
    password: str
    nit: str
    xot_code: str
    got_code: str
    campo_84: str
    campo_05: str
    obl_code: str
    obl2_code: str

    # Timing
    form_ready_wait: float
    login_wait: float
    pyauto_pause: float
    write_interval: float
    wait_default: float
    wait_long: float
    wait_popup: float

    # Inputs de data
    invoices_by_id_path: Path
    fallback_placa: str

    # Robustez / mitigaciones
    stop_on_error: bool
    soft_reset_every: int
    log_every: int
    breath_sec: float
    error_dir: Path
    preflight_icons: bool
    screenshot_on_error: bool

    @staticmethod
    def from_settings() -> "SafixConfig":
        title_re = _strip_quotes(getattr(settings, "safix_window_title", ""))

        def _p(name: str) -> Path:
            v = getattr(settings, name, None)
            if not v:
                raise ValueError(f"Falta settings.{name}")
            return Path(v)

        return SafixConfig(
            jnlp_path=_p("safix_shortcut"),
            main_window_title_re=title_re,
            tesoreria_icon=_p("safix_tesoreria_icon"),
            valores_icon=_p("safix_valores_icon"),
            z_icon=_p("safix_z_icon"),
            engranes_icon=_p("safix_engranes_icon"),
            user=str(getattr(settings, "safix_user", "") or "").strip(),
            password=str(getattr(settings, "safix_pass", "") or "").strip(),
            nit=str(getattr(settings, "safix_nit", "") or "").strip(),
            xot_code=str(getattr(settings, "safix_xot_code", "") or "").strip(),
            got_code=str(getattr(settings, "safix_got_code", "") or "").strip(),
            campo_84=str(getattr(settings, "safix_campo_84", "") or "").strip(),
            campo_05=str(getattr(settings, "safix_campo_05", "") or "").strip(),
            obl_code=str(getattr(settings, "safix_obl_code", "OBL_EXCLU") or "OBL_EXCLU").strip(),
            obl2_code=str(getattr(settings, "safix_obl2_code", "OBL_ANTCON") or "OBL_ANTCON").strip(),
            form_ready_wait=float(getattr(settings, "safix_form_ready_wait", 0) or 0),
            login_wait=float(getattr(settings, "safix_login_wait", 0) or 0),
            pyauto_pause=float(getattr(settings, "safix_pyauto_pause", 0.0) or 0.0),
            write_interval=float(getattr(settings, "safix_write_interval", 0.05) or 0.05),
            wait_default=float(getattr(settings, "safix_wait_default", 0.7) or 0.7),
            wait_long=float(getattr(settings, "safix_wait_long", 1.2) or 1.2),
            wait_popup=float(getattr(settings, "safix_wait_popup", 1.6) or 1.6),
            invoices_by_id_path=_p("invoices_by_id_path"),
            fallback_placa=str(getattr(settings, "safix_placa", "") or "").strip(),
            stop_on_error=bool(getattr(settings, "safix_stop_on_error", False)),
            soft_reset_every=int(getattr(settings, "safix_soft_reset_every", 0) or 0),
            log_every=int(getattr(settings, "safix_log_every", 1) or 1),
            breath_sec=float(getattr(settings, "safix_breath_sec", 0.4) or 0.4),
            error_dir=Path(getattr(settings, "safix_error_dir", "./output/errors")).resolve(),
            preflight_icons=bool(getattr(settings, "safix_preflight_icons", False)),
            screenshot_on_error=bool(getattr(settings, "safix_screenshot_on_error", True)),
        )


# =========================
# Automatizador SAFIX
# =========================
class SafixAutomator:
    def __init__(self, cfg: SafixConfig) -> None:
        self.cfg = cfg
        self._main_handle: Optional[int] = None

        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = self.cfg.pyauto_pause

        self.cfg.error_dir.mkdir(parents=True, exist_ok=True)

    # ---------- Ventanas ----------
    def _wait_for_window_win32(
        self,
        title_re: str,
        timeout: float = 60.0,
        interval: float = 1.0,
    ):
        start = time.time()
        desktop = Desktop(backend="win32")

        while True:
            try:
                win = desktop.window(title_re=title_re)
                if win.exists():
                    return win
            except Exception:
                pass

            if time.time() - start > timeout:
                raise TimeoutError(f"No apareció ventana con título que matchee: {title_re}")
            time.sleep(interval)

    def validate_icon_files(self):
        missing = []
        for p in [
            self.cfg.tesoreria_icon,
            self.cfg.valores_icon,
            self.cfg.z_icon,
            self.cfg.engranes_icon,
        ]:
            if not Path(p).exists():
                missing.append(str(p))

        if missing:
            raise FileNotFoundError(f"Faltan iconos SAFIX: {missing}")

    def launch_and_focus_main(self):
        if not self.cfg.jnlp_path.exists():
            raise FileNotFoundError(f"No se encontró el JNLP: {self.cfg.jnlp_path}")

        os.startfile(str(self.cfg.jnlp_path))

        main_win = self._wait_for_window_win32(self.cfg.main_window_title_re, timeout=90.0)
        self._main_handle = int(main_win.handle)

        app = Application(backend="win32").connect(handle=main_win.handle)
        main_window = app.window(handle=main_win.handle)
        main_window.set_focus()
        return app, main_window

    def refocus_main(self):
        desktop = Desktop(backend="win32")

        if self._main_handle is not None:
            try:
                win = desktop.window(handle=self._main_handle)
                if win.exists():
                    app = Application(backend="win32").connect(handle=win.handle)
                    window = app.window(handle=win.handle)
                    window.set_focus()
                    return app, window
            except Exception:
                pass

        win = self._wait_for_window_win32(self.cfg.main_window_title_re, timeout=20.0)
        self._main_handle = int(win.handle)
        app = Application(backend="win32").connect(handle=win.handle)
        window = app.window(handle=win.handle)
        window.set_focus()
        return app, window

    # ---------- Helpers ----------
    def wait(self, t: Optional[float] = None):
        time.sleep(self.cfg.wait_default if t is None else t)

    def write_text(self, text: str):
        pyautogui.write(str(text), interval=self.cfg.write_interval)

    def write_text_safe(self, text: str, slow_min_interval: float = 0.10):
        """
        Para campos sensibles (login, NIT, document_id, interface, placa, valores):
        selecciona todo, borra, escribe más lento y deja margen.
        """
        pyautogui.hotkey("ctrl", "a")
        self.wait(0.2)
        pyautogui.press("backspace")
        self.wait(0.2)
        pyautogui.write(str(text), interval=max(self.cfg.write_interval, slow_min_interval))
        self.wait(0.6)

    def capture_error_snapshot(self, *, document_id: str = "", placa: str = "", stage: str = "") -> Optional[Path]:
        if not self.cfg.screenshot_on_error:
            return None

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        doc = re.sub(r"[^\w\-]+", "_", document_id or "no_doc")[:50]
        plc = re.sub(r"[^\w\-]+", "_", placa or "no_placa")[:20]
        stg = re.sub(r"[^\w\-]+", "_", stage or "error")[:30]
        name = f"safix_error_{ts}_{doc}_{plc}_{stg}.png"
        out = self.cfg.error_dir / name
        try:
            pyautogui.screenshot(str(out))
            return out
        except Exception:
            return None

    def press_enter(self, n: int = 1, wait_each: Optional[float] = None):
        for _ in range(max(1, n)):
            pyautogui.press("enter")
            if wait_each is not None:
                self.wait(wait_each)

    def press_tab(self, n: int = 1, wait_each: Optional[float] = None):
        for _ in range(max(1, n)):
            pyautogui.press("tab")
            if wait_each is not None:
                self.wait(wait_each)

    # ---------- Login ----------
    def do_login(self):
        self.wait(1.2)
        self.write_text_safe(self.cfg.user)

        pyautogui.press("tab")
        self.wait(0.8)

        self.write_text_safe(self.cfg.password)
        pyautogui.press("enter")

    # ---------- Click imagen ----------
    def click_image(self, image_path: str | Path, timeout: float = 40.0, interval: float = 1.0) -> bool:
        icon_path = Path(image_path)
        if not icon_path.exists():
            raise FileNotFoundError(f"Image not found: {icon_path}")

        start = time.time()
        while time.time() - start < timeout:
            pos = pyautogui.locateCenterOnScreen(str(icon_path))
            if pos:
                pyautogui.click(pos.x, pos.y)
                self.wait(self.cfg.wait_long)
                return True
            time.sleep(interval)

        raise RuntimeError(f"No se pudo ubicar imagen en pantalla: {icon_path}")

    def click_tesoreria(self):
        return self.click_image(self.cfg.tesoreria_icon)

    # ---------- ALT+P,O,G (solo 1 vez en bootstrap) ----------
    def alt_p_o_g(self):
        pyautogui.keyDown("alt")
        time.sleep(0.1)
        pyautogui.press("p")
        time.sleep(0.25)
        pyautogui.press("o")
        time.sleep(0.25)
        pyautogui.press("g")
        time.sleep(0.25)
        pyautogui.keyUp("alt")
        time.sleep(1.2)

    # ---------- Escribir en modal CTRL + L ----------
    def escribir_modal(self, text_str: str):
        pyautogui.hotkey("ctrl", "l")
        self.wait(self.cfg.wait_default)

        self.press_tab(4, self.cfg.wait_default)
        pyautogui.press("right")
        self.wait(self.cfg.wait_default)

        # GOT: write normal (no "safe")
        self.write_text(text_str)
        self.wait(self.cfg.wait_long)

        self.press_enter(2, self.cfg.wait_long)

    # ---------- Proceso por factura ----------
    def procesar_factura(
        self,
        document_id: str,
        placa: str,
        interface: str,
        centro_costos: str,
        peaje_total: int,
        status_cb=None
    ):
        _ = centro_costos  # reservado

        # Si el engrane te devuelve al campo XOT, solo espera a que esté listo
        self.wait(self.cfg.wait_long)

        # XOT
        self.write_text_safe(self.cfg.xot_code)
        pyautogui.press("enter")
        self.wait(self.cfg.wait_long * 1.5)

        # NIT
        self.write_text_safe(self.cfg.nit)
        pyautogui.press("enter")
        self.wait(self.cfg.wait_popup * 1.5)

        # Confirmaciones iniciales
        self.press_enter(1, self.cfg.wait_popup)
        self.press_tab(4, self.cfg.wait_long)

        # Document ID
        self.write_text_safe(document_id)
        pyautogui.press("enter")
        self.wait(self.cfg.wait_long)

        # Confirmaciones posteriores
        self.press_enter(3, self.cfg.wait_long)

        # GOT
        self.escribir_modal(self.cfg.got_code)

        self.press_tab(1, self.cfg.wait_default)

        # Campo 84
        interface_to_write = (interface or "").strip() or self.cfg.campo_84
        self.write_text_safe(interface_to_write)
        self.press_enter(2, self.cfg.wait_long)

        # Campo 05
        self.write_text_safe(self.cfg.campo_05)
        pyautogui.press("enter")
        self.wait(self.cfg.wait_long)
        pyautogui.press("tab")
        self.wait(self.cfg.wait_default)

        # Tabs hasta placa
        self.press_tab(8, self.cfg.wait_default)

        # Placa
        self.write_text_safe(placa)
        self.wait(self.cfg.wait_long * 1.3)

        # Click en valores
        self.click_image(self.cfg.valores_icon)

        # ========= OBL_EXCLU + total(peaje) =========
        self.wait(self.cfg.wait_popup)

        self.write_text_safe(self.cfg.obl_code)
        pyautogui.press("enter")
        self.wait(self.cfg.wait_long)

        self.write_text_safe(str(peaje_total))
        self.press_enter(2, self.cfg.wait_long)
        self.wait(self.cfg.wait_long)

        self.write_text_safe(self.cfg.obl2_code)
        pyautogui.press("enter")
        self.wait(self.cfg.wait_long)

        self.write_text_safe(str(peaje_total))
        self.press_enter(2, self.cfg.wait_long)
        self.wait(self.cfg.wait_long)

        # clic Z
        self.click_image(self.cfg.z_icon)
        self.wait(self.cfg.wait_long)

        self.press_tab(4, self.cfg.wait_long)
        self.escribir_modal(self.cfg.got_code)
        self.wait(self.cfg.wait_long)

        # Cerrar transacción
        self.click_image(self.cfg.engranes_icon)
        self.wait(self.cfg.wait_long)
        self.press_enter(1, self.cfg.wait_long)
        self.wait(self.cfg.wait_default)

        # Limpieza suave para asegurar que quedas listo en XOT
        pyautogui.press("esc")
        self.wait(0.3)
        self.refocus_main()
        self.wait(0.4)

    # ---------- Bootstrap ----------
    def bootstrap(self):
        if self.cfg.preflight_icons:
            self.validate_icon_files()

        self.launch_and_focus_main()
        time.sleep(self.cfg.form_ready_wait)

        self.refocus_main()
        self.wait(1.0)

        self.do_login()
        time.sleep(self.cfg.login_wait)

        self.refocus_main()
        self.wait(1.0)

        self.click_tesoreria()
        self.wait(self.cfg.wait_long * 2)

        # Entrar una sola vez al flujo que deja el cursor en XOT
        self.alt_p_o_g()
        self.wait(self.cfg.wait_long)

    def soft_reset(self):
        """
        Reset suave para recuperar contexto cuando se degrada la UI.
        """
        try:
            self.refocus_main()
            self.wait(0.6)
            self.click_tesoreria()
            self.wait(self.cfg.wait_long * 2)
            self.alt_p_o_g()
            self.wait(self.cfg.wait_long)
        except Exception:
            # si falla, no bloquea el flujo principal
            pass


# =========================
# Runners (BATCH)
# =========================
def run_safix_from_aggregated_json() -> None:
    cfg = SafixConfig.from_settings()

    invoices_by_id = load_invoices_by_id(cfg.invoices_by_id_path)
    if not invoices_by_id:
        print("[SAFIX] No hay facturas. No se ejecuta.")
        _show_no_invoices_message()
        return

    items = list(invoices_by_id.items())
    total_invoices = len(items)

    overlay = StatusOverlay(width=340, height=165)
    overlay.start()
    overlay.update(etapa="AIVO: RENTAN", current=0, total=total_invoices, extra="Iniciando SAFIX…")

    automator = SafixAutomator(cfg)

    print(f"[SAFIX] leyendo agregado: {cfg.invoices_by_id_path.resolve()}")
    print(f"[SAFIX] total facturas: {total_invoices}")

    automator.bootstrap()
    automator.refocus_main()
    automator.wait(1.0)

    ok = 0
    fail = 0

    for idx, (key, invoice) in enumerate(items, start=1):
        doc_id = ""
        placa = ""
        try:
            doc_id, placa, total = extract_doc_plate_and_total(
                invoice,
                fallback_plate=cfg.fallback_placa,
                key_fallback_doc_id=key,
            )

            overlay.update(
                etapa="AIVO: Procesando factura",
                document_id=doc_id,
                placa=placa,
                current=idx,
                total=total_invoices,
                extra="",
            )

            automator.refocus_main()
            automator.wait(0.6)

            automator.procesar_factura(
                document_id=doc_id,
                placa=placa,
                interface="",
                centro_costos="",
                peaje_total=total,
            )

            ok += 1

        except Exception as e:
            fail += 1
            overlay.update(
                etapa="AIVO: ERROR",
                document_id=doc_id,
                placa=placa,
                current=idx,
                total=total_invoices,
                extra=str(e),
            )
            snap = automator.capture_error_snapshot(
                document_id=doc_id,
                placa=placa,
                stage="exception",
            )
            if snap:
                print(f"[SAFIX][ERROR] snapshot: {snap}")
            print(f"[SAFIX][ERROR] key='{key}': {e}")
            try:
                automator.refocus_main()
                automator.wait(1.0)
            except Exception:
                pass
            if cfg.stop_on_error:
                break
            continue

        if cfg.soft_reset_every > 0 and (idx % cfg.soft_reset_every == 0):
            automator.soft_reset()

        if cfg.breath_sec > 0:
            time.sleep(cfg.breath_sec)

    overlay.update(
        etapa="AIVO: Finalizado",
        current=total_invoices,
        total=total_invoices,
        extra=f"OK={ok} FAIL={fail}",
    )
    overlay.stop_timer()
    print(f"[SAFIX] Finalizado. OK={ok} FAIL={fail}")




def run_safix_with_excel(
    excel_path: Path,
    aggregated_json_path: Optional[Path] = None,
    json_root: Optional[Path] = None,
    mark_as_read: bool = False,
    tracker=None,
) -> None:
    """
    Procesa TODAS las facturas del agregado (sin ordenar),
    y por cada una intenta resolver interface/centro_costos por placa desde Excel.

    aggregated_json_path (override):
      Si se pasa, reemplaza cfg.invoices_by_id_path SOLO para esta ejecución
      sin modificar settings/.env (config inmutable + replace()).

    tracker (opcional):
      Instancia de RunTracker para registrar procesadas.xlsx
    """
    from datetime import datetime

    cfg = SafixConfig.from_settings()
    if aggregated_json_path is not None:
        cfg = replace(cfg, invoices_by_id_path=aggregated_json_path)

    plate_catalog = load_plate_catalog_from_excel(excel_path)
    print(f"[SAFIX] catálogo placas cargado: {len(plate_catalog)} desde {excel_path.resolve()}")

    if json_root is not None:
        invoices_by_id = load_invoices_from_json_dir(json_root)
    else:
        invoices_by_id = load_invoices_by_id(cfg.invoices_by_id_path)
    if not invoices_by_id:
        print("[SAFIX] No hay facturas. No se ejecuta.")
        _show_no_invoices_message()
        return

    # Overlay (si lo quieres en esta función)
    overlay = StatusOverlay(width=340, height=165)
    overlay.start()
    overlay.update(etapa="AIVO: RENTAN", extra="Abriendo SAFIX y preparando sesión…")

    # ✅ Deduplicar por document_id
    deduped: Dict[str, Any] = {}
    for key, inv in invoices_by_id.items():
        doc = ((inv.get("document") or {}).get("document_id")) if isinstance(inv, dict) else None
        doc = str(doc or "").strip() or str(key)
        if doc:
            deduped[doc] = inv
    items = list(deduped.items())

    # ✅ Saltar procesadas (si existe procesadas.xlsx)
    processed_ids = set()
    if tracker is not None:
        processed_ids = load_processed_ids_from_log(tracker.logs_dir / "procesadas.xlsx")
    elif json_root is not None:
        processed_ids = load_processed_ids_from_log(Path(settings.output_dir) / "logs" / "procesadas.xlsx")

    if processed_ids:
        items = [(k, v) for (k, v) in items if str(k).strip() not in processed_ids]

    # ✅ preparar marcado como leído (después de procesar cada factura)
    manifest = load_attachment_manifest(settings.download_dir) if mark_as_read else {}
    marked_entry_ids: set[str] = set()
    outlook_client: Optional[OutlookClient] = None
    if mark_as_read and manifest:
        com_ctx = com_initialized()
        com_ctx.__enter__()
        try:
            outlook_client = OutlookClient()
        except Exception:
            com_ctx.__exit__(None, None, None)
            outlook_client = None
            manifest = {}
    else:
        com_ctx = None
    total_invoices = len(items)

    # Inicia contador global desde que aparece overlay:
    overlay.update(
        etapa="AIVO: RENTAN",
        document_id="",
        placa="",
        current=0,
        total=total_invoices,
        extra=f"Facturas detectadas: {total_invoices}",
    )

    automator = SafixAutomator(cfg)

    print(f"[SAFIX] leyendo agregado: {cfg.invoices_by_id_path.resolve()}")
    print(f"[SAFIX] total facturas: {total_invoices}")

    automator.bootstrap()
    automator.refocus_main()
    automator.wait(1.0)

    ok = 0
    fail = 0
    missing_plate = 0

    for idx, (key, invoice) in enumerate(items, start=1):
        doc_id = ""
        placa = ""
        total = None
        interface = ""
        centro_costos = ""

        t_start = datetime.now().isoformat(timespec="seconds")

        try:
            doc_id, placa, total = extract_doc_plate_and_total(
                invoice,
                fallback_plate=cfg.fallback_placa,
                key_fallback_doc_id=key,
            )

            row = plate_catalog.get(placa)
            if not row:
                missing_plate += 1
                interface = ""
                centro_costos = ""
                extra = f"placa no está en catálogo | total={total}"
                if tracker is not None:
                    miss_row = tracker.add_missing_plate(
                        document_id=doc_id,
                        placa=placa,
                        total=int(total) if total is not None else None,
                        key=str(key),
                        error="placa no está en catálogo",
                    )
                    tracker.append_missing_plate_row(miss_row)
            else:
                if str(row.get("doble_cc", "")).strip().upper() == "X":
                    missing_plate += 1
                    extra = f"placa con DOBLE CC | total={total}"
                    overlay.update(
                        etapa="AIVO: OMITIDA",
                        document_id=doc_id,
                        placa=placa,
                        current=idx,
                        total=total_invoices,
                        extra="Omitida por DOBLE CC",
                    )
                    if tracker is not None:
                        miss_row = tracker.add_missing_plate(
                            document_id=doc_id,
                            placa=placa,
                            total=int(total) if total is not None else None,
                            key=str(key),
                            error="placa no procesada por DOBLE CC",
                        )
                        tracker.append_missing_plate_row(miss_row)
                        doble_row = tracker.add_doble_cc(
                            document_id=doc_id,
                            placa=placa,
                            total=int(total) if total is not None else None,
                            key=str(key),
                            error="placa no procesada por DOBLE CC",
                        )
                        tracker.append_doble_cc_row(doble_row)
                    # Saltar procesamiento
                    continue
                interface = row.get("interface", "")
                centro_costos = row.get("centro_costos", "")
                extra = f"interface={interface or '-'} | cc={centro_costos or '-'} | total={total}"

            # ✅ Overlay: mostrar lo que se está procesando en el momento
            overlay.update(
                etapa="AIVO: Procesando factura",
                document_id=doc_id,
                placa=placa,
                current=idx,
                total=total_invoices,
                extra=extra,
            )

            automator.refocus_main()
            automator.wait(0.6)

            automator.procesar_factura(
                document_id=doc_id,
                placa=placa,
                interface=interface,
                centro_costos=centro_costos,
                peaje_total=int(total),
            )

            if mark_as_read and outlook_client is not None:
                src_path = (invoice or {}).get("_source_path") if isinstance(invoice, dict) else None
                zip_stem = _zip_stem_from_source_path(str(src_path)) if src_path else None
                entry_id = manifest.get(zip_stem) if zip_stem else None
                if entry_id and entry_id not in marked_entry_ids:
                    try:
                        outlook_client.mark_as_read_by_id(entry_id)
                        marked_entry_ids.add(entry_id)
                    except Exception:
                        pass

            t_end = datetime.now().isoformat(timespec="seconds")
            if tracker is not None:
                row = tracker.add_processed(
                    timestamp_start=t_start,
                    timestamp_end=t_end,
                    document_id=doc_id,
                    placa=placa,
                    total=int(total) if total is not None else None,
                    key=str(key),
                    status="OK",
                    error="",
                )
                if cfg.log_every <= 1 or (idx % cfg.log_every == 0):
                    tracker.append_processed_row(row)

            ok += 1

        except Exception as e:
            fail += 1
            t_end = datetime.now().isoformat(timespec="seconds")

            # ✅ Overlay: error contextual
            overlay.update(
                etapa="AIVO: ERROR",
                document_id=doc_id,
                placa=placa,
                current=idx,
                total=total_invoices,
                extra=f"{type(e).__name__}: {e}",
            )

            if tracker is not None:
                row = tracker.add_processed(
                    timestamp_start=t_start,
                    timestamp_end=t_end,
                    document_id=doc_id,
                    placa=placa,
                    total=int(total) if isinstance(total, (int, float)) else None,
                    key=str(key),
                    status="FAIL",
                    error=str(e),
                )
                if cfg.log_every <= 1 or (idx % cfg.log_every == 0):
                    tracker.append_processed_row(row)

            snap = automator.capture_error_snapshot(
                document_id=doc_id,
                placa=placa,
                stage="exception",
            )
            if snap:
                print(f"[SAFIX][ERROR] snapshot: {snap}")

            print(f"[SAFIX][ERROR] key='{key}': {e}")
            try:
                automator.refocus_main()
                automator.wait(1.0)
            except Exception:
                pass
            if cfg.stop_on_error:
                break
            continue

        if cfg.soft_reset_every > 0 and (idx % cfg.soft_reset_every == 0):
            automator.soft_reset()

        if cfg.breath_sec > 0:
            time.sleep(cfg.breath_sec)

    overlay.update(
        etapa="AIVO: Finalizado",
        document_id="",
        placa="",
        current=total_invoices,
        total=total_invoices,
        extra=f"OK={ok} FAIL={fail} missing_plate={missing_plate}",
    )
    overlay.stop_timer()

    print(f"[SAFIX] Finalizado. OK={ok} FAIL={fail} missing_plate={missing_plate}")
    # overlay.stop()  # si quieres cerrarlo automáticamente

    if mark_as_read and outlook_client is not None and com_ctx is not None:
        com_ctx.__exit__(None, None, None)


