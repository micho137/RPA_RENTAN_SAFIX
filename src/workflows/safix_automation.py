from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Tuple, Optional

import pyautogui
from pywinauto import Desktop, Application

from src.config import settings


PLACA_REGEX = re.compile(r"Placa:\s*([A-Z0-9]+)", re.IGNORECASE)


def load_invoices_by_id(path: Path) -> Dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"No existe el agregado invoices_by_id: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("El archivo all_invoices_by_id.json no es un dict.")
    return data


def extract_doc_and_plate(invoice: Dict[str, Any], fallback_plate: str) -> Tuple[str, str]:
    doc_id = ((invoice.get("document") or {}).get("document_id")) or ""
    doc_id = str(doc_id).strip()
    if not doc_id:
        raise ValueError("Factura sin document_id (no se puede cargar en SAFIX).")

    placa: Optional[str] = None
    detalles = invoice.get("detalle") or []
    for item in detalles:
        desc = (item or {}).get("descripcion") or ""
        m = PLACA_REGEX.search(desc)
        if m:
            placa = m.group(1).strip()
            break

    if not placa:
        placa = fallback_plate

    return doc_id.replace("-", ""), placa


@dataclass
class SafixConfig:
    jnlp_path: Path
    main_window_title_re: str
    tesoreria_icon: str
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
        return SafixConfig(
            jnlp_path=settings.safix_shortcut,
            main_window_title_re=settings.safix_window_title,
            tesoreria_icon=settings.safix_tesoreria_icon,
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
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = self.cfg.pyauto_pause

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
                raise TimeoutError(f"No apareció ninguna ventana con título: {title_re}")
            time.sleep(interval)

    def launch_and_focus_main(self):
        if not self.cfg.jnlp_path or not self.cfg.jnlp_path.exists():
            raise FileNotFoundError(f"No se encontró el JNLP: {self.cfg.jnlp_path}")
        import os as _os
        _os.startfile(str(self.cfg.jnlp_path))

        main_win = self._wait_for_window_win32(self.cfg.main_window_title_re, timeout=60.0)
        app = Application(backend="win32").connect(handle=main_win.handle)
        main_window = app.window(handle=main_win.handle)
        main_window.set_focus()
        return app, main_window

    def refocus_main(self):
        win = self._wait_for_window_win32(self.cfg.main_window_title_re, timeout=15.0)
        app = Application(backend="win32").connect(handle=win.handle)
        window = app.window(handle=win.handle)
        window.set_focus()
        return app, window

    def wait(self, t: Optional[float] = None):
        time.sleep(self.cfg.wait_default if t is None else t)

    def write_text(self, text: str):
        pyautogui.write(text, interval=self.cfg.write_interval)

    def do_login(self):
        self.write_text(self.cfg.user)
        self.wait()
        pyautogui.press("tab")
        self.wait()
        self.write_text(self.cfg.password)
        self.wait()
        pyautogui.press("enter")

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
        self.wait(self.cfg.wait_long)

    def procesar_factura(self, document_id: str, placa: str):
        self.alt_p_o_g()
        self.wait()

        self.write_text(self.cfg.xot_code)
        self.wait()
        pyautogui.press("enter")
        self.wait(self.cfg.wait_long)

        self.write_text(self.cfg.nit)
        self.wait(self.cfg.wait_popup)
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
        self.write_text(self.cfg.campo_84); self.wait()
        pyautogui.press("enter"); self.wait()
        pyautogui.press("enter"); self.wait()

        self.write_text(self.cfg.campo_05); self.wait()
        pyautogui.press("enter"); self.wait()
        pyautogui.press("tab"); self.wait()

        for _ in range(9):
            pyautogui.press("tab")
        self.wait()

        self.write_text(placa)
        self.wait(self.cfg.wait_long)

    def bootstrap(self):
        self.launch_and_focus_main()
        time.sleep(self.cfg.form_ready_wait)

        self.refocus_main()
        self.wait()

        self.do_login()
        time.sleep(self.cfg.login_wait)

        self.refocus_main()
        self.wait()

        self.click_tesoreria()
        self.wait(self.cfg.wait_long)


def run_safix_from_aggregated_json():
    cfg = SafixConfig.from_settings()

    invoices_by_id = load_invoices_by_id(cfg.invoices_by_id_path)
    if not invoices_by_id:
        print("[SAFIX] No hay facturas. No se ejecuta.")
        return

    automator = SafixAutomator(cfg)
    automator.bootstrap()

    for k in sorted(invoices_by_id.keys()):
        inv = invoices_by_id[k]
        doc_id, placa = extract_doc_and_plate(inv, fallback_plate=cfg.fallback_placa)

        automator.refocus_main()
        automator.wait(0.5)

        automator.procesar_factura(document_id=doc_id, placa=placa)
        automator.wait(cfg.wait_long)

    print("[SAFIX] Carga finalizada.")
