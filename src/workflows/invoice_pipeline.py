from pathlib import Path

from src.config import settings
from src.core.logging_config import setup_logger
from src.workflows.download_attachments import run_download
from src.processing.zip_invoice_extractor import ZipInvoiceExtractor
from src.processing.aggregate_json import build_invoices_by_id
from src.workflows.safix_automation import run_safix_from_aggregated_json


# =========================
# DIRECTORIO RAÍZ DE SALIDA
# =========================

OUTPUT_DIR = Path("./output")

EXTRACT_DIR = OUTPUT_DIR / "extract"
JSON_DIR = OUTPUT_DIR / "json"
TEXT_DIR = OUTPUT_DIR / "text"
INDEX_CSV = JSON_DIR / "index.csv"

# Solo este agregado
AGG_BY_ID_JSON = JSON_DIR / "all_invoices_by_id.json"


def run_pipeline(
    output_dir: Path = OUTPUT_DIR,
    lang: str = "spa",
    dpi: int = 300,
    aggregate_by_id: bool = True,
):
    """
    Pipeline completo:
    - Descarga adjuntos (ZIP) desde Outlook
    - Extrae ZIPs
    - Procesa XML o PDF (OCR si aplica)
    - Genera JSON + TXT
    - Indexa resultados
    - Genera SOLO un JSON agregado por document_id: all_invoices_by_id.json
    Todo queda bajo el directorio `output/`
    """

    # ---------- Preparar directorios ----------
    extract_dir = output_dir / "extract"
    json_dir = output_dir / "json"
    text_dir = output_dir / "text"
    index_csv = json_dir / "index.csv"
    agg_by_id_json = json_dir / "all_invoices_by_id.json"

    extract_dir.mkdir(parents=True, exist_ok=True)
    json_dir.mkdir(parents=True, exist_ok=True)
    text_dir.mkdir(parents=True, exist_ok=True)

    # ---------- Logger ----------
    logger = setup_logger("invoice_pipeline", settings.log_dir)

    # ---------- 1) Descargar adjuntos ----------
    logger.info("Downloading attachments from Outlook...")
    dl = run_download()
    logger.info("Download done: processed=%s | saved=%s", dl.processed, dl.attachments_saved)

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
        overwrite_json=True,  # <- evita duplicados: mismo document_id => mismo archivo
    )

    res = proc.process_all(index_csv=index_csv)
    logger.info("Done. ZIPs=%s | JSONs=%s | Errors=%s", res.zips, res.docs_json, res.errors)

    # ---------- 3) Generar agregado SOLO by_id ----------
    agg_res = None
    if aggregate_by_id:
        try:
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

            # ---------- 4) Ejecutar SAFIX al finalizar ----------
            run_safix_from_aggregated_json()

        except Exception as e:
            logger.exception("Invoices-by-id or SAFIX failed: %s", e)

    return {
        "output_dir": output_dir,
        "download": dl,
        "extract": res,
        "aggregate_by_id": agg_res,
        "paths": {
            "extract": extract_dir,
            "json": json_dir,
            "text": text_dir,
            "index": index_csv,
            "agg_by_id": agg_by_id_json if aggregate_by_id else None,
        },
    }
