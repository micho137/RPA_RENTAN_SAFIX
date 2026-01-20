from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Tuple, Optional

import pyautogui
from openpyxl import load_workbook
from pywinauto import Desktop, Application

from src.config import settings


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

        catalog[placa] = {
            "centro_costos": str(centro_costos or "").strip(),
            "interface": str(interface or "").strip(),
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


def pick_first_invoice(invoices_by_id: Dict[str, Any]) -> Tuple[str, Dict[str, Any]]:
    if not invoices_by_id:
        raise ValueError("No hay facturas en el agregado.")
    first_key = sorted(invoices_by_id.keys())[0]
    return first_key, invoices_by_id[first_key]


# =========================
# Config SAFIX
# =========================
@dataclass
class SafixConfig:
    # App/ventana
    jnlp_path: Path
    main_window_title_re: str

    # UI (icons)
    tesoreria_icon: Path
    valores_icon: Path

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

    @staticmethod
    def from_settings() -> "SafixConfig":
        title_re = _strip_quotes(getattr(settings, "safix_window_title", ""))

        tesoreria_icon = Path(getattr(settings, "safix_tesoreria_icon", "img.png"))
        valores_icon = Path(getattr(settings, "safix_valores_icon", "img_1.png"))

        return SafixConfig(
            jnlp_path=settings.safix_shortcut,
            main_window_title_re=title_re,
            tesoreria_icon=tesoreria_icon,
            valores_icon=valores_icon,
            user=settings.safix_user,
            password=settings.safix_pass,
            nit=str(settings.safix_nit or ""),
            xot_code=settings.safix_xot_code,
            got_code=settings.safix_got_code,
            campo_84=str(settings.safix_campo_84 or ""),
            campo_05=str(settings.safix_campo_05 or ""),
            obl_code=str(getattr(settings, "safix_obl_code", "OBL_EXCLU") or "OBL_EXCLU"),
            obl2_code=str(getattr(settings,"safix_obl2_code","OBL_ANTCON") or "OBL_ANTCON"),
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


# =========================
# Automatizador SAFIX
# =========================
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
                raise TimeoutError(f"No apareció ventana con título que matchee: {title_re}")
            time.sleep(interval)

    def launch_and_focus_main(self):
        if not self.cfg.jnlp_path or not self.cfg.jnlp_path.exists():
            raise FileNotFoundError(f"No se encontró el JNLP: {self.cfg.jnlp_path}")

        import os as _os
        _os.startfile(str(self.cfg.jnlp_path))

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

    def write_text_safe(self, text: str):
        """
        Para campos sensibles (login, NIT, document_id, interface, placa, valores):
        selecciona todo, borra, escribe más lento y deja margen.
        """
        pyautogui.hotkey("ctrl", "a")
        self.wait(0.2)
        pyautogui.press("backspace")
        self.wait(0.2)
        pyautogui.write(str(text), interval=max(self.cfg.write_interval, 0.10))
        self.wait(0.6)

    def press_enter(self, n: int = 1, wait_each: Optional[float] = None):
        for _ in range(max(1, n)):
            pyautogui.press("enter")
            self.wait(wait_each)

    def press_tab(self, n: int = 1, wait_each: Optional[float] = None):
        for _ in range(max(1, n)):
            pyautogui.press("tab")
            self.wait(wait_each)

    # ---------- Login (robusto) ----------
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

    # ---------- ALT+P,O,G ----------
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

    # ---------- Proceso por factura ----------
    def procesar_factura(
        self,
        document_id: str,
        placa: str,
        interface: str,
        centro_costos: str,
        peaje_total: int,  # <-- NUEVO
    ):
        _ = centro_costos  # reservado

        self.alt_p_o_g()
        self.wait(self.cfg.wait_long)

        # XOT
        self.write_text_safe(self.cfg.xot_code)
        pyautogui.press("enter")
        self.wait(self.cfg.wait_long * 2)

        # NIT (punto crítico)
        self.write_text_safe(self.cfg.nit)
        pyautogui.press("enter")
        self.wait(self.cfg.wait_popup * 2)

        # Confirmaciones iniciales
        self.press_enter(1, self.cfg.wait_popup)
        self.press_enter(4, self.cfg.wait_long)

        # Document ID (no saltar)
        self.write_text_safe(document_id)
        pyautogui.press("enter")
        self.wait(self.cfg.wait_long)

        # Confirmaciones posteriores
        self.press_enter(3, self.cfg.wait_long)

        # Navegación
        pyautogui.hotkey("ctrl", "l")
        self.wait(self.cfg.wait_default)
        self.press_tab(4, self.cfg.wait_default)
        pyautogui.press("right")
        self.wait(self.cfg.wait_default)

        # GOT (IMPORTANTE: SOLO write_text, NO write_text_safe)
        self.write_text(self.cfg.got_code)
        self.wait(self.cfg.wait_long)

        self.press_tab(2, self.cfg.wait_default)
        self.press_enter(2, self.cfg.wait_long)

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
        self.wait(self.cfg.wait_long * 2)

        # Click en valores
        self.click_image(self.cfg.valores_icon)

        # ========= NUEVO: OBL_EXCLU + total(peaje) =========
        self.wait(self.cfg.wait_popup)

        # escribir OBL_EXCLU (desde .env)
        self.write_text_safe(self.cfg.obl_code)
        pyautogui.press("enter")
        self.wait(self.cfg.wait_long)

        # escribir total
        self.write_text_safe(str(peaje_total))
        self.press_enter(2, self.cfg.wait_long)
        self.wait(self.cfg.wait_long)

        # escribir OBL_ANTCON
        self.write_text_safe(self.cfg.obl2_code)
        pyautogui.press("enter")
        self.wait(self.cfg.wait_long)

        # escribir total
        self.write_text_safe(str(peaje_total))
        self.press_enter(2, self.cfg.wait_long)
        self.wait(self.cfg.wait_long)

        # hacer clic en imagen Z


    # ---------- Bootstrap ----------
    def bootstrap(self):
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


# =========================
# Runners
# =========================
def run_safix_from_aggregated_json() -> None:
    cfg = SafixConfig.from_settings()

    invoices_by_id = load_invoices_by_id(cfg.invoices_by_id_path)
    if not invoices_by_id:
        print("[SAFIX] No hay facturas. No se ejecuta.")
        return

    automator = SafixAutomator(cfg)

    print(f"[SAFIX] leyendo agregado: {cfg.invoices_by_id_path.resolve()}")
    print(f"[SAFIX] total facturas: {len(invoices_by_id)}")

    first_key, first_invoice = pick_first_invoice(invoices_by_id)

    doc_id, placa, total = extract_doc_plate_and_total(
        first_invoice,
        fallback_plate=cfg.fallback_placa,
        key_fallback_doc_id=first_key,
    )

    print(f"[SAFIX] DEMO → doc_id='{doc_id}' placa='{placa}' total='{total}' key='{first_key}'")

    automator.bootstrap()
    automator.refocus_main()
    automator.wait(1.0)

    automator.procesar_factura(
        document_id=doc_id,
        placa=placa,
        interface="",
        centro_costos="",
        peaje_total=total,
    )

    print("[SAFIX] DEMO finalizada (una sola factura).")


def run_safix_with_excel(excel_path: Path, aggregated_json_path: Optional[Path] = None) -> None:
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

    first_key, first_invoice = pick_first_invoice(invoices_by_id)

    doc_id, placa, total = extract_doc_plate_and_total(
        first_invoice,
        fallback_plate=cfg.fallback_placa,
        key_fallback_doc_id=first_key,
    )

    row = plate_catalog.get(placa, {})
    interface = row.get("interface", "")
    centro_costos = row.get("centro_costos", "")

    print(
        f"[SAFIX] DEMO → doc_id='{doc_id}' placa='{placa}' "
        f"interface='{interface}' centro_costos='{centro_costos}' total='{total}'"
    )

    automator.bootstrap()
    automator.refocus_main()
    automator.wait(1.0)

    automator.procesar_factura(
        document_id=doc_id,
        placa=placa,
        interface=interface,
        centro_costos=centro_costos,
        peaje_total=total,
    )

    print("[SAFIX] DEMO finalizada (una sola factura) con Excel.")
