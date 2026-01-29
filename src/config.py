from __future__ import annotations

import os
import sys
from dataclasses import dataclass
from pathlib import Path
from dotenv import load_dotenv


# =========================
# Resolución base (DEV vs EXE)
# =========================
def _base_dir() -> Path:
    """
    - En ejecutable (.exe): carpeta donde vive el ejecutable
    - En desarrollo: raíz del proyecto
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


BASE_DIR = _base_dir()
ENV_PATH = BASE_DIR / ".env"

# Cargar .env explícitamente desde el directorio base
if ENV_PATH.exists():
    load_dotenv(dotenv_path=ENV_PATH)
else:
    raise RuntimeError(f"No se encontró el archivo .env en {ENV_PATH}")


# =========================
# Helper obligatorio
# =========================
def _env(name: str) -> str:
    val = os.getenv(name)
    if val is None or val.strip() == "":
        raise RuntimeError(f"Variable de entorno obligatoria no definida: {name}")
    return val.strip()


def _env_path(name: str) -> Path:
    return Path(_env(name)).expanduser().resolve()


def _env_float(name: str) -> float:
    try:
        return float(_env(name))
    except ValueError:
        raise RuntimeError(f"Variable {name} debe ser numérica")


# =========================
# Settings
# =========================
@dataclass(frozen=True)
class Settings:
    # ---------- Outlook ----------
    outlook_account: str
    source_folder: str
    download_dir: Path
    processed_folder: str
    log_dir: Path

    # ---------- Pipeline ----------
    output_dir: Path
    extract_dir: Path
    json_dir: Path
    text_dir: Path
    index_csv: Path
    invoices_by_id_path: Path

    # ---------- SAFIX ----------
    safix_shortcut: Path
    safix_window_title: str

    safix_tesoreria_icon: Path
    safix_valores_icon: Path
    safix_z_icon: Path
    safix_engranes_icon: Path

    safix_user: str
    safix_pass: str
    safix_nit: str

    safix_xot_code: str
    safix_got_code: str
    safix_campo_84: str
    safix_campo_05: str
    safix_placa: str

    safix_obl_code: str
    safix_obl2_code: str

    safix_form_ready_wait: float
    safix_login_wait: float
    safix_pyauto_pause: float
    safix_write_interval: float
    safix_wait_default: float
    safix_wait_long: float
    safix_wait_popup: float


# =========================
# Construcción explícita
# =========================
settings = Settings(
    # Outlook
    outlook_account=_env("OUTLOOK_ACCOUNT"),
    source_folder=_env("OUTLOOK_FOLDER"),
    download_dir=_env_path("DOWNLOAD_DIR"),
    processed_folder=_env("PROCESSED_FOLDER"),
    log_dir=_env_path("LOG_DIR"),

    # Pipeline
    output_dir=_env_path("OUTPUT_DIR"),
    extract_dir=_env_path("EXTRACT_DIR"),
    json_dir=_env_path("JSON_DIR"),
    text_dir=_env_path("TEXT_DIR"),
    index_csv=_env_path("INDEX_CSV"),
    invoices_by_id_path=_env_path("INVOICES_BY_ID_PATH"),

    # SAFIX
    safix_shortcut=_env_path("SAFIX_SHORTCUT"),
    safix_window_title=_env("SAFIX_WINDOW_TITLE"),

    safix_tesoreria_icon=_env_path("SAFIX_TESORERIA_ICON"),
    safix_valores_icon=_env_path("SAFIX_VALORES_ICON"),
    safix_z_icon=_env_path("SAFIX_Z_ICON"),
    safix_engranes_icon=_env_path("SAFIX_ENGRANES_ICON"),

    safix_user=_env("SAFIX_USER"),
    safix_pass=_env("SAFIX_PASS"),
    safix_nit=_env("SAFIX_NIT"),

    safix_xot_code=_env("SAFIX_XOT_CODE"),
    safix_got_code=_env("SAFIX_GOT_CODE"),
    safix_campo_84=_env("SAFIX_CAMPO_84"),
    safix_campo_05=_env("SAFIX_CAMPO_05"),
    safix_placa=_env("SAFIX_PLACA"),

    safix_obl_code=_env("SAFIX_OBL_CODE"),
    safix_obl2_code=_env("SAFIX_OBL2_CODE"),

    safix_form_ready_wait=_env_float("SAFIX_FORM_READY_WAIT"),
    safix_login_wait=_env_float("SAFIX_LOGIN_WAIT"),
    safix_pyauto_pause=_env_float("SAFIX_PYAUTO_PAUSE"),
    safix_write_interval=_env_float("SAFIX_WRITE_INTERVAL"),
    safix_wait_default=_env_float("SAFIX_WAIT_DEFAULT"),
    safix_wait_long=_env_float("SAFIX_WAIT_LONG"),
    safix_wait_popup=_env_float("SAFIX_WAIT_POPUP"),
)


# =========================
# Crear directorios
# =========================
for p in [
    settings.download_dir,
    settings.log_dir,
    settings.output_dir,
    settings.extract_dir,
    settings.json_dir,
    settings.text_dir,
]:
    p.mkdir(parents=True, exist_ok=True)
