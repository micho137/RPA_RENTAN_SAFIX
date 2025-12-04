from pathlib import Path
from src.config import settings
from src.core.logging_config import setup_logger
from src.workflows.download_attachments import run_download
from src.processing.zip_invoice_extractor import ZipInvoiceExtractor

BASE_OUTPUT = Path("./output")

DEFAULT_EXTRACT = BASE_OUTPUT / "unzipped"
DEFAULT_JSON   = BASE_OUTPUT / "json"
DEFAULT_TEXT   = BASE_OUTPUT / "text"      # para los .txt de PDF/OCR
DEFAULT_INDEX  = BASE_OUTPUT / "index.csv"

def run_pipeline(
    extract_dir: Path = DEFAULT_EXTRACT,
    json_dir: Path = DEFAULT_JSON,
    text_dir: Path = DEFAULT_TEXT,
    index_csv: Path = DEFAULT_INDEX,
    lang: str = "spa",
    dpi: int = 300,
):
    logger = setup_logger("invoice_pipeline", settings.log_dir)

    # Crear las carpetas si no existen
    extract_dir.mkdir(parents=True, exist_ok=True)
    json_dir.mkdir(parents=True, exist_ok=True)
    text_dir.mkdir(parents=True, exist_ok=True)
    index_csv.parent.mkdir(parents=True, exist_ok=True)

    logger.info("Downloading attachments from Outlook...")
    dl = run_download()
    logger.info(f"Download done: processed={dl.processed} | saved={dl.attachments_saved}")

    logger.info("Extracting and processing ZIPs...")
    proc = ZipInvoiceExtractor(
        download_dir=settings.download_dir,
        extract_root=extract_dir,
        out_json_root=json_dir,
        out_text_root=text_dir,
        lang=lang,
        dpi=dpi,
        logger=logger,
    )
    res = proc.process_all(index_csv=index_csv)
    logger.info(f"Done. ZIPs={res.zips} | JSONs={res.docs_json} | Errors={res.errors}")
    return {"download": dl, "extract": res}
