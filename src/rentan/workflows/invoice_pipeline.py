from __future__ import annotations

import logging
from pathlib import Path

from src.rentan.config.config import settings
from src.rentan.workflows.download_attachments import run_download
from src.rentan.processing.zip_invoice_extractor import ZipInvoiceExtractor
from src.rentan.processing.aggregate_json import build_invoices_by_id
from src.rentan.workflows.safix_automation import run_safix_with_excel

from src.rentan.core.run_tracking import RunTracker
from src.rentan.core.cleanup import cleanup_output_dir_keep_pdf_xml


logger = logging.getLogger(__name__)

OUTPUT_DIR = Path("./output")


def run_pipeline(
    output_dir: Path = OUTPUT_DIR,
    lang: str = "spa",
    dpi: int = 300,
    aggregate_by_id: bool = True,
    excel_path: Path | None = None,
    run_safix: bool = True,
    only_unread: bool = False,
    days_back: int = 0,
    mark_as_read: bool = False,
    move_to_processed: bool = False,
):
    """
    Pipeline completo:
    - Descarga adjuntos (ZIP) desde Outlook
    - Extrae ZIPs
    - Procesa XML o PDF (OCR si aplica)
    - Genera JSON + TXT
    - Indexa resultados
    - Genera agregado por document_id (requerido por SAFIX en este flujo)
    - Ejecuta SAFIX (opcional) con Excel
    - Registra procesadas.xlsx de forma incremental (desde SAFIX) o al final si tu tracker lo hace así
    - Cleanup final: elimina todo excepto PDF/XML y el log Excel de procesadas
    """

    output_dir = Path(output_dir).resolve()

    extract_dir = output_dir / "extract"
    json_dir = output_dir / "json"
    text_dir = output_dir / "text"
    index_csv = json_dir / "index.csv"
    agg_by_id_json = json_dir / "all_invoices_by_id.json"

    extract_dir.mkdir(parents=True, exist_ok=True)
    json_dir.mkdir(parents=True, exist_ok=True)
    text_dir.mkdir(parents=True, exist_ok=True)

    logger.info("Pipeline start | output_dir=%s", output_dir)
    logger.info(
        "[PIPE][FLAGS] only_unread=%s days_back=%s mark_as_read=%s move_to_processed=%s run_safix=%s aggregate_by_id=%s",
        only_unread,
        days_back,
        mark_as_read,
        move_to_processed,
        run_safix,
        aggregate_by_id,
    )

    tracker = RunTracker(output_dir=output_dir)

    # 1) Descargar adjuntos
    logger.info("Downloading attachments from Outlook...")
    dl = run_download(
        days_back=days_back,
        only_unread=only_unread,
        mark_as_read=mark_as_read,
        move_to_processed=move_to_processed,
    )
    logger.info("Download done | processed=%s | saved=%s | errors=%s", dl.processed, dl.attachments_saved, dl.errors)

    # 2) Procesar ZIPs
    logger.info("Extracting and processing ZIPs...")
    proc = ZipInvoiceExtractor(
        download_dir=settings.download_dir,
        extract_root=extract_dir,
        out_json_root=json_dir,
        out_text_root=text_dir,
        lang=lang,
        dpi=dpi,
        logger=logger,
        overwrite_json=True,
    )

    ex = proc.process_all(index_csv=index_csv)
    logger.info("Extract done | zips=%s | docs_json=%s | errors=%s", ex.zips, ex.docs_json, ex.errors)

    # 3) Agregado por ID (lo mantiene el flujo actual porque SAFIX lo consume)
    agg_res = None
    if aggregate_by_id:
        agg_res = build_invoices_by_id(
            json_root=json_dir,
            out_by_id_path=agg_by_id_json,
            include_source_path=True,
        )
        logger.info(
            "Invoices-by-id generated | total_ids=%s | path=%s",
            agg_res.total_ids,
            agg_res.written_by_id,
        )

    # 4) SAFIX
    if run_safix:
        if excel_path is None:
            raise ValueError("run_safix=True pero excel_path=None. Debes seleccionar el Excel desde Flet o CLI.")
        excel_path = Path(excel_path).resolve()
        if not excel_path.exists():
            raise FileNotFoundError(f"El Excel no existe: {excel_path}")

        if not agg_by_id_json.exists():
            raise FileNotFoundError(f"No se generó el agregado esperado: {agg_by_id_json}")

        logger.info("Running SAFIX | excel=%s | aggregated_json=%s", excel_path, agg_by_id_json)

        run_safix_with_excel(
            excel_path=excel_path,
            aggregated_json_path=agg_by_id_json,
            tracker=tracker,
        )

    # tracker.save() si lo mantienes por compatibilidad, pero tu nueva intención es incremental.
    # Igual lo invocamos si existe y no rompe. Si tu RunTracker ya escribe incrementalmente,
    # esto no debe “duplicar” si está bien diseñado.
    paths = {}
    try:
        paths = tracker.save()
        procesadas_path = paths.get("procesadas")
    except Exception as e:
        logger.warning("Tracker save failed: %s", e)
        procesadas_path = None

    if procesadas_path:
        logger.info("Logs generated | procesadas=%s", procesadas_path)

    # Cleanup: conservar PDFs, XMLs y procesadas.xlsx (si existe)
    keep_paths = tuple(p for p in [procesadas_path] if p is not None)
    deleted_files, deleted_dirs = cleanup_output_dir_keep_pdf_xml(
        output_dir=output_dir,
        keep_exts=(".pdf", ".xml"),
        keep_paths=keep_paths,
    )
    logger.info("[CLEANUP] deleted_files=%s deleted_dirs=%s", deleted_files, deleted_dirs)

    return {
        "output_dir": output_dir,
        "download": dl,
        "extract": ex,
        "aggregate_by_id": agg_res,
        "logs": {"procesadas": procesadas_path},
        "cleanup": {"deleted_files": deleted_files, "deleted_dirs": deleted_dirs},
        "paths": {
            "extract": extract_dir,
            "json": json_dir,
            "text": text_dir,
            "index": index_csv,
            "agg_by_id": agg_by_id_json if aggregate_by_id else None,
        },
    }
