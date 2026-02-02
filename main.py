from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

import flet as ft

# ======================================================
# Asegurar imports sin PYTHONPATH
# ======================================================
# main.py está fuera de src/, así que añadimos src/ al path
BASE_DIR = Path(__file__).resolve().parent
SRC_DIR = BASE_DIR / "src"

if SRC_DIR.exists():
    sys.path.insert(0, str(SRC_DIR))
else:
    raise RuntimeError("No se encontró la carpeta 'src'. Estructura inválida.")

# ======================================================
# Imports del proyecto
# ======================================================
from src.rentan.config.config import settings, ensure_dirs
from src.rentan.core.logging_config import (
    setup_logging,
    setup_stdout_stderr_to_logging,
    install_global_excepthook,
)

from src.rentan.workflows.download_attachments import run_download
from src.rentan.workflows.invoice_pipeline import run_pipeline


# ======================================================
# UI (Flet)
# ======================================================
def ejecutar_ui() -> None:
    """
    Lanza la interfaz gráfica de Flet.
    Uso:
        python main.py --ui
    """
    setup_logging(settings.log_dir)
    setup_stdout_stderr_to_logging()
    install_global_excepthook()

    ensure_dirs(settings)

    logger = logging.getLogger("main")
    logger.info("Iniciando interfaz gráfica (Flet)")

    from src.rentan.ui.ui_flet import main as ui_main

    ft.app(
        target=ui_main,
        view=ft.AppView.FLET_APP,
    )


# ======================================================
# CLI
# ======================================================
def main() -> None:
    parser = argparse.ArgumentParser(
        description="RENTAN - Automatización Outlook + Pipeline + SAFIX"
    )

    grupo = parser.add_mutually_exclusive_group()
    grupo.add_argument("--ui", action="store_true", help="Abrir interfaz gráfica (Flet)")
    grupo.add_argument("--download", action="store_true", help="Solo descargar adjuntos desde Outlook")
    grupo.add_argument("--pipeline", action="store_true", help="Ejecutar pipeline completo")

    # Flags Outlook
    parser.add_argument("--days-back", type=int, default=0, help="Días hacia atrás (0 = todos)")
    parser.add_argument("--only-unread", action="store_true", help="Solo correos no leídos")
    parser.add_argument("--mark-as-read", action="store_true", help="Marcar correos como leídos")
    parser.add_argument("--move-to-processed", action="store_true", help="Mover correos a carpeta Procesados")

    # Pipeline
    parser.add_argument("--output", type=str, default=str(settings.output_dir), help="Directorio de salida")
    parser.add_argument("--lang", default="spa", help="Idioma OCR")
    parser.add_argument("--dpi", type=int, default=300, help="DPI OCR")
    parser.add_argument("--excel", type=str, help="Ruta al Excel de Centros de Costos")

    parser.add_argument("--no-safix", action="store_true", help="No ejecutar SAFIX")
    parser.add_argument("--no-aggregate", action="store_true", help="No generar agregado por ID")

    args = parser.parse_args()

    # Inicialización global
    setup_logging(settings.log_dir)
    setup_stdout_stderr_to_logging()
    install_global_excepthook()
    ensure_dirs(settings)

    logger = logging.getLogger("main")
    logger.info("Parámetros recibidos: %s", vars(args))

    # ---------------- UI ----------------
    if args.ui:
        ejecutar_ui()
        return

    # ---------------- DOWNLOAD ----------------
    if args.download:
        res = run_download(
            days_back=args.days_back,
            only_unread=args.only_unread,
            mark_as_read=args.mark_as_read,
            move_to_processed=args.move_to_processed,
        )
        logger.info(
            "Descarga finalizada | procesados=%s | guardados=%s",
            res.processed,
            res.attachments_saved,
        )
        print(f"Procesados={res.processed} | Guardados={res.attachments_saved}")
        return

    # ---------------- PIPELINE ----------------
    if args.pipeline:
        if not args.excel and not args.no_safix:
            raise RuntimeError("Debe especificar --excel si SAFIX está habilitado.")

        resultado = run_pipeline(
            output_dir=Path(args.output).resolve(),
            lang=args.lang,
            dpi=args.dpi,
            aggregate_by_id=not args.no_aggregate,
            excel_path=Path(args.excel).resolve() if args.excel else None,
            run_safix=not args.no_safix,
            only_unread=args.only_unread,
            days_back=args.days_back,
            mark_as_read=args.mark_as_read,
            move_to_processed=args.move_to_processed,
        )

        logger.info("Pipeline finalizado correctamente")
        print(f"Salida: {resultado['output_dir']}")
        return

    parser.print_help()


if __name__ == "__main__":
    main()
