# src/workflows/invoice_pipeline.py
from pathlib import Path

from src.config import settings
from src.core.logging_config import setup_logger
from src.workflows.download_attachments import run_download
from src.processing.zip_invoice_extractor import ZipInvoiceExtractor
from src.processing.aggregate_json import build_invoices_by_id
from src.workflows.safix_automation import run_safix_with_excel

from src.core.run_tracking import RunTracker
from src.core.cleanup import cleanup_output_dir_keep_pdf_xml


OUTPUT_DIR = Path("./output")
AGG_BY_ID_JSON = (OUTPUT_DIR / "json" / "all_invoices_by_id.json")


def run_pipeline(
    output_dir: Path = OUTPUT_DIR,
    lang: str = "spa",
    dpi: int = 300,
    aggregate_by_id: bool = True,
    excel_path: Path | None = None,
    run_safix: bool = True,

    # ✅ flags dinámicos (UI)
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
    - Genera SOLO un JSON agregado por document_id: all_invoices_by_id.json
    - (Opcional) Ejecuta SAFIX al finalizar usando Excel para INTERFACE/CENTRO DE COSTOS
    - (Nuevo) Genera logs Excel: descargados.xlsx y procesadas.xlsx
    - (Nuevo) Cleanup final: elimina todo excepto PDF y XML (y conserva los logs Excel)
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

    logger = setup_logger("invoice_pipeline", settings.log_dir)

    # ✅ Tracker para Excels
    tracker = RunTracker(output_dir=output_dir)

    # ---------- 1) Descargar adjuntos ----------
    logger.info("Downloading attachments from Outlook...")
    logger.info(
        "[PIPE][FLAGS] only_unread=%s days_back=%s mark_as_read=%s move_to_processed=%s",
        only_unread, days_back, mark_as_read, move_to_processed
    )

    dl = run_download(
        days_back=days_back,
        only_unread=only_unread,
        mark_as_read=mark_as_read,
        move_to_processed=move_to_processed,
    )
    logger.info("Download done: processed=%s | saved=%s", dl.processed, dl.attachments_saved)

    # ✅ Log descargados (ZIPs)
    # Asumimos que settings.download_dir es el dir donde se guardan los adjuntos.
    # Y que dl.attachments_saved es un contador. Para listar exactamente qué se guardó,
    # hacemos un snapshot: los ZIP en settings.download_dir.
    try:
        download_dir = Path(settings.download_dir).resolve()
        if download_dir.exists():
            for p in sorted(download_dir.glob("*.zip")):
                tracker.add_download(
                    original_name=p.name,
                    saved_path=p,
                    size_bytes=p.stat().st_size if p.exists() else None,
                    status="OK",
                    error="",
                )
    except Exception as ex:
        logger.warning("No se pudo registrar descargados.xlsx (ZIP listing). Detalle: %s", ex)

    # ---------- 2) Procesar ZIPs ----------
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

    res = proc.process_all(index_csv=index_csv)
    logger.info("Done. ZIPs=%s | JSONs=%s | Errors=%s", res.zips, res.docs_json, res.errors)

    # ---------- 3) Generar agregado SOLO by_id ----------
    agg_res = None
    if aggregate_by_id:
        agg_res = build_invoices_by_id(
            json_root=json_dir,
            out_by_id_path=agg_by_id_json,
            include_source_path=True,
        )
        logger.info(
            "Invoices-by-id generated: total_ids=%s | path=%s",
            agg_res.total_ids,
            agg_res.written_by_id,
        )

        # ---------- 4) SAFIX ----------
        if run_safix:
            if excel_path is None:
                raise ValueError("run_safix=True pero excel_path=None. Debes seleccionar el Excel desde Flet.")

            excel_path = Path(excel_path).resolve()
            if not excel_path.exists():
                raise FileNotFoundError(f"El Excel no existe: {excel_path}")

            if not agg_by_id_json.exists():
                raise FileNotFoundError(f"No se generó el agregado esperado: {agg_by_id_json}")

            logger.info("Running SAFIX with Excel=%s and AggregatedJSON=%s", excel_path, agg_by_id_json)

            run_safix_with_excel(
                excel_path=excel_path,
                aggregated_json_path=agg_by_id_json,
                tracker=tracker,  # ✅ registra procesadas.xlsx
            )

    # ✅ Guardar Excels
    paths = tracker.save()
    logger.info("Logs generated: descargados=%s | procesadas=%s", paths["descargados"], paths["procesadas"])

    # ✅ Cleanup final: conservar PDFs, XMLs y los logs Excel
    deleted_files, deleted_dirs = cleanup_output_dir_keep_pdf_xml(
        output_dir=output_dir,
        keep_exts=(".pdf", ".xml"),
        keep_paths=(paths["descargados"], paths["procesadas"]),
    )
    logger.info("[CLEANUP] deleted_files=%s deleted_dirs=%s", deleted_files, deleted_dirs)

    return {
        "output_dir": output_dir,
        "download": dl,
        "extract": res,
        "aggregate_by_id": agg_res,
        "logs": {"descargados": paths["descargados"], "procesadas": paths["procesadas"]},
        "cleanup": {"deleted_files": deleted_files, "deleted_dirs": deleted_dirs},
        "paths": {
            "extract": extract_dir,
            "json": json_dir,
            "text": text_dir,
            "index": index_csv,
            "agg_by_id": agg_by_id_json if aggregate_by_id else None,
        },
    }
