from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Tuple, Optional

import pyautogui
from pywinauto import Desktop, Application

from openpyxl import load_workbook

from src.config import settings


PLACA_REGEX = re.compile(
    r"\bPLACA\b\s*[:\-]?\s*([A-Z0-9]{5,8})\b",
    re.IGNORECASE
)

def _strip_quotes(s: str) -> str:
    """Quita comillas simples/dobles envolventes si existen."""
    if not s:
        return s
    s = s.strip()
    if (s.startswith("'") and s.endswith("'")) or (s.startswith('"') and s.endswith('"')):
        return s[1:-1].strip()
    return s


def load_invoices_by_id(path: Path) -> Dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"No existe el agregado invoices_by_id: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("El archivo all_invoices_by_id.json no es un dict.")
    return data


def normalize_placa(s: str) -> str:
    """Normaliza placa a mayúsculas y sin espacios."""
    return re.sub(r"\s+", "", str(s or "").strip().upper())


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
    ws = wb.active  # si está en otra hoja, cámbialo por wb["NombreHoja"]

    # Leer encabezados (primera fila)
    headers: Dict[str, int] = {}
    for col_idx, cell in enumerate(ws[1], start=1):
        h = str(cell.value or "").strip().upper()
        if h:
            headers[h] = col_idx

    required = ["PLACA", "CENTRO DE COSTOS", "INTERFACE"]
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

        centro_costos_s = str(centro_costos or "").strip()
        interface_s = str(interface or "").strip()

        # Si hay duplicados de placa, la última fila sobreescribe
        catalog[placa] = {
            "centro_costos": centro_costos_s,
            "interface": interface_s,
        }

    return catalog


def extract_doc_and_plate(
    invoice: Dict[str, Any],
    fallback_plate: str,
    key_fallback_doc_id: Optional[str] = None,
) -> Tuple[str, str]:
    # -------- document_id --------
    doc_id = ((invoice.get("document") or {}).get("document_id")) or ""
    doc_id = str(doc_id).strip()

    # respaldo: la key del agregado (tu all_invoices_by_id.json está indexado por eso)
    if not doc_id and key_fallback_doc_id:
        doc_id = str(key_fallback_doc_id).strip()

    if not doc_id:
        raise ValueError("Factura sin document_id (ni en JSON ni en la key del agregado).")

    doc_id = doc_id.replace("-", "").strip()

    # -------- placa --------
    placa = (
        invoice.get("placa")
        or (invoice.get("vehiculo") or {}).get("placa")
        or ""
    )
    placa = str(placa).strip()

    if not placa:
        detalles = invoice.get("detalle") or []
        for item in detalles:
            desc = str((item or {}).get("descripcion") or "")
            m = PLACA_REGEX.search(desc)
            if m:
                placa = m.group(1).strip()
                break

    if not placa:
        # si el env dice "PLACA HYT426", extraemos solo HYT426
        m2 = PLACA_REGEX.search(str(fallback_plate))
        placa = (m2.group(1).strip() if m2 else str(fallback_plate).strip())

    return doc_id, placa


@dataclass
class SafixConfig:
    jnlp_path: Path
    main_window_title_re: str
    tesoreria_icon: str
    valores_icon: str
    user: str
    password: str

    nit: str
    xot_code: str
    got_code: str
    campo_84: str
    campo_05: str

    form_ready_wait: float
    login_wait: float
    pyauto_pause: float
    write_interval: float
    wait_default: float
    wait_long: float
    wait_popup: float

    invoices_by_id_path: Path
    fallback_placa: str

    @staticmethod
    def from_settings() -> "SafixConfig":
        # IMPORTANTE: quitar comillas del regex si vienen desde .env
        title_re = _strip_quotes(getattr(settings, "safix_window_title", ""))

        return SafixConfig(
            jnlp_path=settings.safix_shortcut,
            main_window_title_re=title_re,
            tesoreria_icon="src/img.png",
            valores_icon="src/valores.png",
            user=settings.safix_user,
            password=settings.safix_pass,
            nit=settings.safix_nit,
            xot_code=settings.safix_xot_code,
            got_code=settings.safix_got_code,
            campo_84=settings.safix_campo_84,
            campo_05=settings.safix_campo_05,
            form_ready_wait=settings.safix_form_ready_wait,
            login_wait=settings.safix_login_wait,
            pyauto_pause=settings.safix_pyauto_pause,
            write_interval=settings.safix_write_interval,
            wait_default=settings.safix_wait_default,
            wait_long=settings.safix_wait_long,
            wait_popup=settings.safix_wait_popup,
            invoices_by_id_path=settings.invoices_by_id_path,
            fallback_placa=settings.safix_placa,
        )


class SafixAutomator:
    def __init__(self, cfg: SafixConfig) -> None:
        self.cfg = cfg
        self._main_handle: Optional[int] = None

        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = self.cfg.pyauto_pause

    # ---------- Ventanas ----------
    def _wait_for_window_win32(self, title_re: str, timeout: float = 60.0, interval: float = 1.0):
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
                raise TimeoutError(f"No apareció ninguna ventana con título que matchee: {title_re}")
            time.sleep(interval)

    def launch_and_focus_main(self):
        if not self.cfg.jnlp_path or not self.cfg.jnlp_path.exists():
            raise FileNotFoundError(f"No se encontró el JNLP: {self.cfg.jnlp_path}")

        import os as _os
        _os.startfile(str(self.cfg.jnlp_path))

        main_win = self._wait_for_window_win32(self.cfg.main_window_title_re, timeout=60.0)
        self._main_handle = int(main_win.handle)  # guardar handle para reenfocar aunque cambie título

        app = Application(backend="win32").connect(handle=main_win.handle)
        main_window = app.window(handle=main_win.handle)
        main_window.set_focus()
        return app, main_window

    def refocus_main(self):
        """
        Reenfoca por handle (robusto si cambia el título dentro de Tesorería).
        Si no hay handle, cae a búsqueda por título.
        """
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

        # Fallback: por título
        win = self._wait_for_window_win32(self.cfg.main_window_title_re, timeout=15.0)
        self._main_handle = int(win.handle)
        app = Application(backend="win32").connect(handle=win.handle)
        window = app.window(handle=win.handle)
        window.set_focus()
        return app, window

    # ---------- Helpers ----------
    def wait(self, t: Optional[float] = None):
        time.sleep(self.cfg.wait_default if t is None else t)

    def write_text(self, text: str):
        pyautogui.write(text, interval=self.cfg.write_interval)

    # ---------- Login ----------
    def do_login(self):
        time.sleep(0.8)
        self.write_text(self.cfg.user)
        time.sleep(0.6)
        pyautogui.press("tab")
        time.sleep(0.6)
        self.write_text(self.cfg.password)
        time.sleep(0.6)
        pyautogui.press("enter")

    # ---------- Tesorería ----------
    def click_tesoreria(self, timeout: float = 30.0, interval: float = 1.0):
        icon_path = Path(self.cfg.tesoreria_icon)
        if not icon_path.exists():
            raise FileNotFoundError(f"No se encontró el icono de Tesorería: {icon_path}")

        start = time.time()
        while time.time() - start < timeout:
            pos = pyautogui.locateCenterOnScreen(str(icon_path))
            if pos:
                pyautogui.click(pos.x, pos.y)
                self.wait(self.cfg.wait_long)
                return True
            time.sleep(interval)

        raise RuntimeError("No se pudo encontrar el icono de Tesorería en pantalla.")

    # ---------- ALT + P, O, G ----------
    def alt_p_o_g(self):
        pyautogui.keyDown("alt")
        time.sleep(0.1)
        pyautogui.press("p")
        time.sleep(0.2)
        pyautogui.press("o")
        time.sleep(0.2)
        pyautogui.press("g")
        time.sleep(0.2)
        pyautogui.keyUp("alt")
        time.sleep(1.0)

    # ---------- Proceso por factura ----------
    def procesar_factura(self, document_id: str, placa: str, interface: str, centro_costos: str):
        # centro_costos queda disponible (no se usa aún, por solicitud)
        _ = centro_costos

        self.alt_p_o_g()
        self.wait()

        self.write_text(self.cfg.xot_code)
        self.wait()
        pyautogui.press("enter")
        self.wait(self.cfg.wait_long)

        self.write_text(self.cfg.nit)
        time.sleep(2.0)
        pyautogui.press("enter")
        self.wait(self.cfg.wait_popup)

        pyautogui.press("enter")
        self.wait(self.cfg.wait_long)

        pyautogui.press("enter"); self.wait()
        pyautogui.press("enter"); self.wait()
        pyautogui.press("enter"); self.wait()
        pyautogui.press("enter"); self.wait()

        self.write_text(document_id)
        self.wait()
        pyautogui.press("enter"); self.wait()
        pyautogui.press("enter"); self.wait()
        pyautogui.press("enter"); self.wait()
        pyautogui.press("enter"); self.wait()

        pyautogui.hotkey("ctrl", "l"); self.wait()
        for _ in range(4):
            pyautogui.press("tab")
        self.wait()
        pyautogui.press("right"); self.wait()

        self.write_text(self.cfg.got_code)
        self.wait()

        pyautogui.press("tab"); pyautogui.press("tab")
        self.wait()
        pyautogui.press("enter"); self.wait()
        pyautogui.press("enter"); self.wait()

        pyautogui.press("tab"); self.wait()

        # REEMPLAZO: campo_84 se llena con INTERFACE por placa, con fallback a cfg.campo_84
        interface_to_write = (interface or "").strip() or self.cfg.campo_84
        self.write_text(interface_to_write); self.wait()

        pyautogui.press("enter"); self.wait()
        pyautogui.press("enter"); self.wait()

        self.write_text(self.cfg.campo_05); self.wait()
        pyautogui.press("enter"); self.wait()
        pyautogui.press("tab"); self.wait()

        for _ in range(8):
            pyautogui.press("tab")
        self.wait()

        self.write_text(placa)
        self.wait(self.cfg.wait_long)

    # ---------- Bootstrap ----------
    def bootstrap(self):
        self.launch_and_focus_main()
        time.sleep(self.cfg.form_ready_wait)

        self.refocus_main()
        time.sleep(0.5)

        self.do_login()
        time.sleep(self.cfg.login_wait)

        self.refocus_main()
        time.sleep(0.5)

        self.click_tesoreria()
        self.wait(self.cfg.wait_long)


def run_safix_from_aggregated_json():
    """
    Runner original (sin Excel). Se mantiene por compatibilidad.
    """
    cfg = SafixConfig.from_settings()

    invoices_by_id = load_invoices_by_id(cfg.invoices_by_id_path)
    if not invoices_by_id:
        print("[SAFIX] No hay facturas. No se ejecuta.")
        return

    automator = SafixAutomator(cfg)

    print(f"[SAFIX] leyendo agregado: {cfg.invoices_by_id_path.resolve()}")
    print(f"[SAFIX] total facturas: {len(invoices_by_id)}")

    first_key = sorted(invoices_by_id.keys())[1]
    first_invoice = invoices_by_id[first_key]

    doc_id, placa = extract_doc_and_plate(
        first_invoice,
        fallback_plate=cfg.fallback_placa,
        key_fallback_doc_id=first_key,
    )

    print(f"[SAFIX] DEMO → doc_id='{doc_id}' | placa='{placa}' | key='{first_key}'")

    automator.bootstrap()
    automator.refocus_main()
    time.sleep(0.5)

    automator.procesar_factura(
        document_id=doc_id,
        placa=placa,
        interface="",
        centro_costos="",
    )

    print("[SAFIX] DEMO finalizada (una sola factura).")


def run_safix_with_excel(excel_path: Path, aggregated_json_path: Optional[Path] = None) -> None:
    """
    Ejecuta SAFIX tomando INTERFACE/CENTRO DE COSTOS desde el Excel seleccionado en Flet.

    - excel_path: archivo con columnas PLACA, CENTRO DE COSTOS, INTERFACE
    - aggregated_json_path: opcional, si quieres sobreescribir cfg.invoices_by_id_path
    """
    cfg = SafixConfig.from_settings()

    if aggregated_json_path is not None:
        cfg.invoices_by_id_path = aggregated_json_path

    plate_catalog = load_plate_catalog_from_excel(excel_path)
    print(f"[SAFIX] catálogo placas cargado: {len(plate_catalog)} desde {excel_path.resolve()}")

    invoices_by_id = load_invoices_by_id(cfg.invoices_by_id_path)
    if not invoices_by_id:
        print("[SAFIX] No hay facturas. No se ejecuta.")
        return

    automator = SafixAutomator(cfg)

    print(f"[SAFIX] leyendo agregado: {cfg.invoices_by_id_path.resolve()}")
    print(f"[SAFIX] total facturas: {len(invoices_by_id)}")

    # DEMO: primera factura (mantengo tu comportamiento)
    first_key = sorted(invoices_by_id.keys())[1]
    first_invoice = invoices_by_id[first_key]

    doc_id, placa = extract_doc_and_plate(
        first_invoice,
        fallback_plate=cfg.fallback_placa,
        key_fallback_doc_id=first_key,
    )

    placa_norm = normalize_placa(placa)
    row = plate_catalog.get(placa_norm, {})
    interface = row.get("interface", "")
    centro_costos = row.get("centro_costos", "")

    print(
        f"[SAFIX] DEMO → doc_id='{doc_id}' "
        f"placa='{placa_norm}' interface='{interface}' centro_costos='{centro_costos}'"
    )

    automator.bootstrap()
    automator.refocus_main()
    time.sleep(0.5)

    automator.procesar_factura(
        document_id=doc_id,
        placa=placa_norm,
        interface=interface,
        centro_costos=centro_costos,
    )

    print("[SAFIX] DEMO finalizada (una sola factura) con Excel.")
