from pathlib import Path

from src.config import settings
from src.core.logging_config import setup_logger
from src.workflows.download_attachments import run_download
from src.processing.zip_invoice_extractor import ZipInvoiceExtractor
from src.processing.aggregate_json import build_invoices_by_id

# CAMBIO: usar el runner que recibe excel_path + aggregated_json_path
from src.workflows.safix_automation import run_safix_with_excel


# =========================
# DIRECTORIO RAÍZ DE SALIDA
# =========================
OUTPUT_DIR = Path("./output")

EXTRACT_DIR = OUTPUT_DIR / "extract"
JSON_DIR = OUTPUT_DIR / "json"
TEXT_DIR = OUTPUT_DIR / "text"
INDEX_CSV = JSON_DIR / "index.csv"

AGG_BY_ID_JSON = JSON_DIR / "all_invoices_by_id.json"


def run_pipeline(
    output_dir: Path = OUTPUT_DIR,
    lang: str = "spa",
    dpi: int = 300,
    aggregate_by_id: bool = True,
    excel_path: Path | None = None,          # NUEVO
    run_safix: bool = True,                   # NUEVO
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
    """

    # Normalizar output_dir (evita problemas por "working directory" distinto)
    output_dir = Path(output_dir).resolve()

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
        overwrite_json=True,
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

            # ---------- 4) Ejecutar SAFIX (si aplica) ----------
            if run_safix:
                if excel_path is None:
                    raise ValueError("run_safix=True pero excel_path=None. Debes seleccionar el Excel desde Flet.")

                excel_path = Path(excel_path).resolve()
                if not excel_path.exists():
                    raise FileNotFoundError(f"El Excel no existe: {excel_path}")

                if not agg_by_id_json.exists():
                    raise FileNotFoundError(f"No se generó el agregado esperado: {agg_by_id_json}")

                logger.info("Running SAFIX with Excel=%s and AggregatedJSON=%s", excel_path, agg_by_id_json)

                # Esto ejecuta SAFIX usando el agregado generado y reemplaza campo_84 por INTERFACE según placa
                run_safix_with_excel(
                    excel_path=excel_path,
                    aggregated_json_path=agg_by_id_json,
                )

        except Exception as e:
            logger.exception("Invoices-by-id or SAFIX failed: %s", e)
            raise  # importante: que Flet capture el error

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
