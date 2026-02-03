from dataclasses import dataclass
import os
from datetime import datetime
from dotenv import load_dotenv
from pathlib import Path

load_dotenv()


def _env_bool(name: str, default: str = "false") -> bool:
    return os.getenv(name, default).strip().lower() in {"1", "true", "yes", "y", "si", "sí"}


def _dated_dir(base: str) -> Path:
    base_path = Path(base).resolve()
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    return base_path / stamp


@dataclass(frozen=True)
class Settings:
    # =========================
    # OUTLOOK / DESCARGA (estático)
    # =========================
    outlook_account: str = os.getenv("OUTLOOK_ACCOUNT", "")
    source_folder: str = os.getenv("OUTLOOK_FOLDER", "")

    download_dir: Path = Path(os.getenv("DOWNLOAD_DIR", "./downloads")).resolve()
    processed_folder: str = os.getenv("PROCESSED_FOLDER", "Procesados")
    log_dir: Path = Path(os.getenv("LOG_DIR", "./src/logs")).resolve()

    # =========================
    # PIPELINE / OUTPUT (estático)
    # =========================
    output_dir: Path = _dated_dir(os.getenv("OUTPUT_DIR", "./output"))

    extract_dir: Path = (
        Path(os.getenv("EXTRACT_DIR", "")).resolve()
        if os.getenv("EXTRACT_DIR")
        else (output_dir / "extract")
    )
    json_dir: Path = (
        Path(os.getenv("JSON_DIR", "")).resolve()
        if os.getenv("JSON_DIR")
        else (output_dir / "json")
    )
    text_dir: Path = (
        Path(os.getenv("TEXT_DIR", "")).resolve()
        if os.getenv("TEXT_DIR")
        else (output_dir / "text")
    )

    index_csv: Path = (
        Path(os.getenv("INDEX_CSV", "")).resolve()
        if os.getenv("INDEX_CSV")
        else (json_dir / "index.csv")
    )
    invoices_by_id_path: Path = (
        Path(os.getenv("INVOICES_BY_ID_PATH", "")).resolve()
        if os.getenv("INVOICES_BY_ID_PATH")
        else (json_dir / "all_invoices_by_id.json")
    )

    # =========================
    # SAFIX / XENCO (estático)
    # =========================
    safix_shortcut: Path = (
        Path(os.getenv("SAFIX_SHORTCUT", "")).resolve()
        if os.getenv("SAFIX_SHORTCUT")
        else Path("")
    )

    safix_window_title: str = os.getenv("SAFIX_WINDOW_TITLE", r".*XENCO - Administracion del Sistema.*")

    safix_tesoreria_icon: str = os.getenv("SAFIX_TESORERIA_ICON", "assets/img.png")
    safix_valores_icon: str = os.getenv("SAFIX_VALORES_ICON", "assets/img_1.png")
    safix_z_icon: str = os.getenv("SAFIX_Z_ICON", "assets/Z.png")
    safix_engranes_icon: str = os.getenv("SAFIX_ENGRANES_ICON", "assets/engranes.png")

    safix_user: str = os.getenv("SAFIX_USER", "")
    safix_pass: str = os.getenv("SAFIX_PASS", "")

    safix_nit: str = os.getenv("SAFIX_NIT", "")
    safix_xot_code: str = os.getenv("SAFIX_XOT_CODE", "XOT05")
    safix_got_code: str = os.getenv("SAFIX_GOT_CODE", "GOT")

    safix_campo_84: str = os.getenv("SAFIX_CAMPO_84", "84")
    safix_campo_05: str = os.getenv("SAFIX_CAMPO_05", "05")
    safix_placa: str = os.getenv("SAFIX_PLACA")

    safix_obl_code: str = os.getenv("SAFIX_OBL_CODE", "OBL_EXCLU")
    safix_obl2_code: str = os.getenv("SAFIX_OBL2_CODE", "OBL_ANTCON")

    safix_form_ready_wait: float = float(os.getenv("SAFIX_FORM_READY_WAIT", "50.0"))
    safix_login_wait: float = float(os.getenv("SAFIX_LOGIN_WAIT", "15.0"))
    safix_pyauto_pause: float = float(os.getenv("SAFIX_PYAUTO_PAUSE", "1.1"))
    safix_write_interval: float = float(os.getenv("SAFIX_WRITE_INTERVAL", "0.10"))
    safix_wait_default: float = float(os.getenv("SAFIX_WAIT_DEFAULT", "1.5"))
    safix_wait_long: float = float(os.getenv("SAFIX_WAIT_LONG", "3.0"))
    safix_wait_popup: float = float(os.getenv("SAFIX_WAIT_POPUP", "4.0"))

    # =========================
    # SAFIX / ROBUSTEZ (dinámico)
    # =========================
    safix_stop_on_error: bool = _env_bool("SAFIX_STOP_ON_ERROR", "false")
    safix_soft_reset_every: int = int(os.getenv("SAFIX_SOFT_RESET_EVERY", "0") or 0)
    safix_log_every: int = int(os.getenv("SAFIX_LOG_EVERY", "1") or 1)
    safix_breath_sec: float = float(os.getenv("SAFIX_BREATH_SEC", "0.4") or 0.4)
    safix_error_dir: Path = Path(os.getenv("SAFIX_ERROR_DIR", "./output/errors")).resolve()
    safix_preflight_icons: bool = _env_bool("SAFIX_PREFLIGHT_ICONS", "false")
    safix_screenshot_on_error: bool = _env_bool("SAFIX_SCREENSHOT_ON_ERROR", "true")


settings = Settings()

# Dirs
settings.download_dir.mkdir(parents=True, exist_ok=True)
settings.log_dir.mkdir(parents=True, exist_ok=True)

settings.output_dir.mkdir(parents=True, exist_ok=True)
settings.extract_dir.mkdir(parents=True, exist_ok=True)
settings.json_dir.mkdir(parents=True, exist_ok=True)
settings.text_dir.mkdir(parents=True, exist_ok=True)
settings.safix_error_dir.mkdir(parents=True, exist_ok=True)
