from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional, Dict, Any

from src.rentan.workflows.download_attachments import run_download
from src.rentan.processing.zip_invoice_extractor import ZipInvoiceExtractor
from src.rentan.processing.aggregate_json import build_invoices_by_id
from src.rentan.workflows.safix_automation import run_safix_with_excel
from src.rentan.core.run_tracking import RunTracker

logger = logging.getLogger(__name__)


def run_pipeline(
    *,
    output_dir: Path,
    lang: str = "spa",
    dpi: int = 300,
    aggregate_by_id: bool = True,
    excel_path: Optional[Path] = None,
    run_safix: bool = True,
    only_unread: bool = False,
    days_back: int = 0,
    mark_as_read: bool = False,
    move_to_processed: bool = False,
) -> Dict[str, Any]:
    """
    Pipeline principal:
      1) Descarga adjuntos de Outlook
      2) Extrae ZIPs
      3) Genera JSONs y agregado por ID
      4) (Opcional) Ejecuta SAFIX

    Regla clave:
      - Si NO hay facturas (agregado vacío) → NO se ejecuta SAFIX
    """

    output_dir = Path(output_dir).resolve()
    extract_dir = output_dir / "extract"
    json_dir = output_dir / "json"
    text_dir = output_dir / "text"

    extract_dir.mkdir(parents=True, exist_ok=True)
    json_dir.mkdir(parents=True, exist_ok=True)
    text_dir.mkdir(parents=True, exist_ok=True)

    logger.info("Pipeline start | output_dir=%s", output_dir)

    tracker = RunTracker(output_dir=output_dir)

    # ======================================================
    # 1) DESCARGA
    # ======================================================
    logger.info(
        "[PIPE][FLAGS] only_unread=%s days_back=%s mark_as_read=%s move_to_processed=%s run_safix=%s aggregate_by_id=%s",
        only_unread,
        days_back,
        mark_as_read,
        move_to_processed,
        run_safix,
        aggregate_by_id,
    )

    dl_result = run_download(
        only_unread=only_unread,
        days_back=days_back,
        mark_as_read=mark_as_read,
        move_to_processed=move_to_processed,
    )

    logger.info(
        "Download done | processed=%s | saved=%s | errors=%s",
        dl_result.processed,
        dl_result.attachments_saved,
        dl_result.errors,
    )

    # ======================================================
    # 2) EXTRACCIÓN ZIPs
    # ======================================================
    logger.info("Extracting and processing ZIPs...")

    extractor = ZipInvoiceExtractor(
        extract_dir=extract_dir,
        json_dir=json_dir,
        text_dir=text_dir,
        lang=lang,
        dpi=dpi,
        tracker=tracker,
    )

    extract_result = extractor.run()

    logger.info(
        "Extract done | zips=%s | docs_json=%s | errors=%s",
        extract_result.zips,
        extract_result.docs_json,
        extract_result.errors,
    )

    # ======================================================
    # 3) AGREGADO POR ID
    # ======================================================
    aggregated = {}
    aggregated_path = None

    if aggregate_by_id:
        aggregated_path = json_dir / "all_invoices_by_id.json"
        aggregated = build_invoices_by_id(
            json_dir=json_dir,
            output_path=aggregated_path,
        )

        logger.info(
            "Invoices-by-id generated | total_ids=%s | path=%s",
            len(aggregated),
            aggregated_path.resolve(),
        )

    # ======================================================
    # 4) VALIDACIÓN CRÍTICA: ¿HAY FACTURAS?
    # ======================================================
    if not aggregated:
        logger.warning(
            "No hay facturas agregadas. SAFIX NO será ejecutado."
        )

        tracker.save()

        return {
            "output_dir": output_dir,
            "extract": extract_result,
            "aggregated_count": 0,
            "safix_executed": False,
            "reason": "No hay facturas para procesar",
        }

    # ======================================================
    # 5) SAFIX (solo si hay facturas)
    # ======================================================
    if run_safix:
        if not excel_path:
            raise ValueError(
                "Se solicitó ejecutar SAFIX pero no se proporcionó excel_path"
            )

        logger.info(
            "Running SAFIX | excel=%s | aggregated_json=%s",
            excel_path,
            aggregated_path,
        )

        run_safix_with_excel(
            excel_path=excel_path,
            aggregated_json_path=aggregated_path,
            tracker=tracker,
        )
    else:
        logger.info("SAFIX deshabilitado por flags.")

    # ======================================================
    # 6) GUARDAR TRACKING
    # ======================================================
    tracker.save()

    logger.info("Pipeline finished successfully.")

    return {
        "output_dir": output_dir,
        "extract": extract_result,
        "aggregated_count": len(aggregated),
        "safix_executed": bool(run_safix),
    }
