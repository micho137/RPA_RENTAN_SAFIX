from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv


def _require(name: str) -> str:
    v = os.getenv(name)
    if v is None or v.strip() == "":
        raise RuntimeError(f"Missing required env var: {name}")
    return v


def _opt(name: str, default: str = "") -> str:
    v = os.getenv(name)
    return default if v is None else v


@dataclass(frozen=True)
class Settings:
    # Outlook
    outlook_account: str
    source_folder: str
    processed_folder: str

    download_dir: Path
    log_dir: Path

    # Output/pipeline
    output_dir: Path
    extract_dir: Path
    json_dir: Path
    text_dir: Path
    index_csv: Path
    invoices_by_id_path: Path

    # Safix (credenciales requeridas)
    safix_user: str
    safix_pass: str
    safix_nit: str

    # Safix (aplicación / ventana)
    safix_shortcut: Path
    safix_window_title: str

    # Safix (icons / UI)
    safix_tesoreria_icon: Path
    safix_valores_icon: Path
    safix_z_icon: Path
    safix_engranes_icon: Path

    # Safix otros
    safix_xot_code: str
    safix_got_code: str
    safix_campo_84: str
    safix_campo_05: str
    safix_obl_code: str
    safix_obl2_code: str

    # Timing
    safix_form_ready_wait: float
    safix_login_wait: float
    safix_pyauto_pause: float
    safix_write_interval: float
    safix_wait_default: float
    safix_wait_long: float
    safix_wait_popup: float


def load_settings(*, env_file: Optional[Path] = None) -> Settings:
    # Siempre permitir .env externo (usuario final)
    if env_file is None:
        env_file = Path(".env")
    if env_file.exists():
        load_dotenv(dotenv_path=env_file)

    # Rutas con defaults OK
    download_dir = Path(_opt("DOWNLOAD_DIR", "./downloads")).resolve()
    log_dir = Path(_opt("LOG_DIR", "./logs")).resolve()

    output_dir = Path(_opt("OUTPUT_DIR", "./output")).resolve()
    extract_dir = Path(_opt("EXTRACT_DIR", str(output_dir / "extract"))).resolve()
    json_dir = Path(_opt("JSON_DIR", str(output_dir / "json"))).resolve()
    text_dir = Path(_opt("TEXT_DIR", str(output_dir / "text"))).resolve()

    index_csv = Path(_opt("INDEX_CSV", str(json_dir / "index.csv"))).resolve()
    invoices_by_id_path = Path(
        _opt("INVOICES_BY_ID_PATH", str(json_dir / "all_invoices_by_id.json"))
    ).resolve()

    # Íconos: por defecto en src/rentan/assets/...
    tesoreria_icon = Path(_opt("SAFIX_TESORERIA_ICON", "src/rentan/assets/tesoreria.png")).resolve()
    valores_icon = Path(_opt("SAFIX_VALORES_ICON", "src/rentan/assets/valores.png")).resolve()
    z_icon = Path(_opt("SAFIX_Z_ICON", "src/rentan/assets/Z.png")).resolve()
    engranes_icon = Path(_opt("SAFIX_ENGRANES_ICON", "src/rentan/assets/engranes.png")).resolve()

    return Settings(
        outlook_account=_require("OUTLOOK_ACCOUNT"),
        source_folder=_require("OUTLOOK_FOLDER"),
        processed_folder=_opt("PROCESSED_FOLDER", "Procesados"),

        download_dir=download_dir,
        log_dir=log_dir,

        output_dir=output_dir,
        extract_dir=extract_dir,
        json_dir=json_dir,
        text_dir=text_dir,
        index_csv=index_csv,
        invoices_by_id_path=invoices_by_id_path,

        # Credenciales SIEMPRE requeridas
        safix_user=_require("SAFIX_USER"),
        safix_pass=_require("SAFIX_PASS"),
        safix_nit=_require("SAFIX_NIT"),

        safix_shortcut=Path(_require("SAFIX_SHORTCUT")).resolve(),
        safix_window_title=_opt("SAFIX_WINDOW_TITLE", r".*XENCO - Administracion del Sistema.*"),

        # Icons / UI
        safix_tesoreria_icon=tesoreria_icon,
        safix_valores_icon=valores_icon,
        safix_z_icon=z_icon,
        safix_engranes_icon=engranes_icon,

        safix_xot_code=_opt("SAFIX_XOT_CODE", "XOT05"),
        safix_got_code=_opt("SAFIX_GOT_CODE", "GOT"),

        safix_campo_84=_opt("SAFIX_CAMPO_84", "84"),
        safix_campo_05=_opt("SAFIX_CAMPO_05", "05"),

        safix_obl_code=_opt("SAFIX_OBL_CODE", "OBL_EXCLU"),
        safix_obl2_code=_opt("SAFIX_OBL2_CODE", "OBL_ANTCON"),

        safix_form_ready_wait=float(_opt("SAFIX_FORM_READY_WAIT", "50.0")),
        safix_login_wait=float(_opt("SAFIX_LOGIN_WAIT", "15.0")),
        safix_pyauto_pause=float(_opt("SAFIX_PYAUTO_PAUSE", "1.1")),
        safix_write_interval=float(_opt("SAFIX_WRITE_INTERVAL", "0.10")),
        safix_wait_default=float(_opt("SAFIX_WAIT_DEFAULT", "1.5")),
        safix_wait_long=float(_opt("SAFIX_WAIT_LONG", "3.0")),
        safix_wait_popup=float(_opt("SAFIX_WAIT_POPUP", "4.0")),
    )


def ensure_dirs(settings: Settings) -> None:
    settings.download_dir.mkdir(parents=True, exist_ok=True)
    settings.log_dir.mkdir(parents=True, exist_ok=True)

    settings.output_dir.mkdir(parents=True, exist_ok=True)
    settings.extract_dir.mkdir(parents=True, exist_ok=True)
    settings.json_dir.mkdir(parents=True, exist_ok=True)
    settings.text_dir.mkdir(parents=True, exist_ok=True)


settings = load_settings()
