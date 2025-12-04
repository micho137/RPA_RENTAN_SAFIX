import argparse
from pathlib import Path

from src.workflows.download_attachments import run_download
from src.workflows.invoice_pipeline import run_pipeline, DEFAULT_EXTRACT, DEFAULT_JSON, DEFAULT_TEXT, DEFAULT_INDEX

def main():
    parser = argparse.ArgumentParser(description="Outlook COM automation CLI")
    parser.add_argument("--download", action="store_true", help="Download attachments using .env settings")

    # Nuevo: ejecutar el pipeline (descarga + extracción XML-first / PDF→OCR)
    parser.add_argument("--pipeline", action="store_true", help="Download + process ZIPs to JSON")

    # Opcionales para el pipeline (si no se pasan, usa defaults)
    parser.add_argument("--extract", default=str(DEFAULT_EXTRACT), help="Dir for unzipped content")
    parser.add_argument("--json",    default=str(DEFAULT_JSON),    help="Dir for JSON outputs")
    parser.add_argument("--text",    default=str(DEFAULT_TEXT),    help="Dir for extracted/OCR text")
    parser.add_argument("--index",   default=str(DEFAULT_INDEX),   help="CSV index path")
    parser.add_argument("--lang",    default="spa",                help="OCR language when needed")
    parser.add_argument("--dpi",     type=int, default=300,        help="OCR DPI (fallback)")

    args = parser.parse_args()

    if args.download and args.pipeline:
        parser.error("Use solo una opción: --download o --pipeline")

    if args.download:
        res = run_download()
        print(f"Processed: {res.processed} | Attachments saved: {res.attachments_saved}")
        return

    if args.pipeline:
        result = run_pipeline(
            extract_dir=Path(args.extract),
            json_dir=Path(args.json),
            text_dir=Path(args.text),
            index_csv=Path(args.index),
            lang=args.lang,
            dpi=args.dpi,
        )
        ex = result["extract"]
        print(f"Pipeline -> ZIPs: {ex.zips} | JSONs: {ex.docs_json} | Errors: {ex.errors}")
        return

    parser.print_help()

if __name__ == "__main__":
    main()
