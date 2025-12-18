from dataclasses import dataclass
import os
from dotenv import load_dotenv
from pathlib import Path

load_dotenv()


@dataclass(frozen=True)
class Settings:
    # =========================
    # OUTLOOK / DESCARGA (EXISTENTE)
    # =========================
    outlook_account: str = os.getenv("OUTLOOK_ACCOUNT", "")
    source_folder: str = os.getenv("OUTLOOK_FOLDER", "")

    download_dir: Path = Path(os.getenv("DOWNLOAD_DIR", "./downloads")).resolve()

    only_unread: bool = os.getenv("ONLY_UNREAD", "false").lower() == "true"
    days_back: int = int(os.getenv("DAYS_BACK", "30"))
    mark_as_read: bool = os.getenv("MARK_AS_READ", "true").lower() == "true"
    move_to_processed: bool = os.getenv("MOVE_TO_PROCESSED", "false").lower() == "true"
    processed_folder: str = os.getenv("PROCESSED_FOLDER", "Procesados")

    log_dir: Path = Path(os.getenv("LOG_DIR", "./logs")).resolve()

    # =========================
    # PIPELINE / OUTPUT (NUEVO, SIN ROMPER NADA)
    # =========================
    output_dir: Path = Path(os.getenv("OUTPUT_DIR", "./output")).resolve()

    extract_dir: Path = Path(os.getenv("EXTRACT_DIR", "")).resolve() if os.getenv("EXTRACT_DIR") else (output_dir / "extract")
    json_dir: Path = Path(os.getenv("JSON_DIR", "")).resolve() if os.getenv("JSON_DIR") else (output_dir / "json")
    text_dir: Path = Path(os.getenv("TEXT_DIR", "")).resolve() if os.getenv("TEXT_DIR") else (output_dir / "text")

    index_csv: Path = Path(os.getenv("INDEX_CSV", "")).resolve() if os.getenv("INDEX_CSV") else (json_dir / "index.csv")
    invoices_by_id_path: Path = Path(os.getenv("INVOICES_BY_ID_PATH", "")).resolve() if os.getenv("INVOICES_BY_ID_PATH") else (json_dir / "all_invoices_by_id.json")

    # OCR (si lo quieres parametrizable desde env)
    ocr_lang: str = os.getenv("OCR_LANG", "spa")
    ocr_dpi: int = int(os.getenv("OCR_DPI", "300"))

    # =========================
    # SAFIX / XENCO (NUEVO)
    # =========================
    safix_shortcut: Path = Path(os.getenv("SAFIX_SHORTCUT", "")).resolve() if os.getenv("SAFIX_SHORTCUT") else Path("")
    safix_window_title: str = os.getenv("SAFIX_WINDOW_TITLE", r".*XENCO - Administracion del Sistema.*").strip("'").strip('"')
    safix_tesoreria_icon: str = os.getenv("SAFIX_TESORERIA_ICON", "img.png")

    safix_user: str = os.getenv("SAFIX_USER", "")
    safix_pass: str = os.getenv("SAFIX_PASS", "")

    safix_nit: str = os.getenv("SAFIX_NIT", "")
    safix_xot_code: str = os.getenv("SAFIX_XOT_CODE", "XOT05")
    safix_got_code: str = os.getenv("SAFIX_GOT_CODE", "GOT")

    safix_campo_84: str = os.getenv("SAFIX_CAMPO_84", "84")
    safix_campo_05: str = os.getenv("SAFIX_CAMPO_05", "05")

    safix_placa: str = os.getenv("SAFIX_PLACA", "PLACA DEFAULT")

    # Timing / estabilidad pyautogui
    safix_form_ready_wait: float = float(os.getenv("SAFIX_FORM_READY_WAIT", "25.0"))
    safix_login_wait: float = float(os.getenv("SAFIX_LOGIN_WAIT", "5.0"))
    safix_pyauto_pause: float = float(os.getenv("SAFIX_PYAUTO_PAUSE", "0.55"))
    safix_write_interval: float = float(os.getenv("SAFIX_WRITE_INTERVAL", "0.03"))
    safix_wait_default: float = float(os.getenv("SAFIX_WAIT_DEFAULT", "0.6"))
    safix_wait_long: float = float(os.getenv("SAFIX_WAIT_LONG", "1.0"))
    safix_wait_popup: float = float(os.getenv("SAFIX_WAIT_POPUP", "1.5"))


settings = Settings()

# Mantener exactamente tu comportamiento actual + agregar los nuevos dirs
settings.download_dir.mkdir(parents=True, exist_ok=True)
settings.log_dir.mkdir(parents=True, exist_ok=True)

settings.output_dir.mkdir(parents=True, exist_ok=True)
settings.extract_dir.mkdir(parents=True, exist_ok=True)
settings.json_dir.mkdir(parents=True, exist_ok=True)
settings.text_dir.mkdir(parents=True, exist_ok=True)
