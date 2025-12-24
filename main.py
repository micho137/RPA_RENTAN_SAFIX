import argparse
from pathlib import Path

from src.workflows.download_attachments import run_download
from src.workflows.invoice_pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="Outlook COM automation CLI")

    group = parser.add_mutually_exclusive_group()
    group.add_argument("--download", action="store_true", help="Download attachments using .env settings")
    group.add_argument("--pipeline", action="store_true", help="Download + process ZIPs to JSON/TXT (output/)")

    # Nuevo: un solo directorio raíz
    parser.add_argument(
        "--output",
        type=str,
        default="./output",
        help="Root output directory (extract/, json/, text/)",
    )

    # Opcionales del pipeline
    parser.add_argument("--lang", default="spa", help="OCR language when needed")
    parser.add_argument("--dpi", type=int, default=300, help="OCR DPI (fallback)")

    args = parser.parse_args()

    if args.download:
        res = run_download()
        print(f"Processed: {res.processed} | Attachments saved: {res.attachments_saved}")
        return

    if args.pipeline:
        result = run_pipeline(
            output_dir=Path(args.output),
            lang=args.lang,
            dpi=args.dpi,
        )
        ex = result["extract"]
        out = result["output_dir"]
        idx = result["paths"]["index"]
        print(f"Output: {out.resolve()}")
        print(f"Index:  {idx.resolve()}")
        print(f"Pipeline -> ZIPs: {ex.zips} | JSONs: {ex.docs_json} | Errors: {ex.errors}")
        return

    parser.print_help()


if __name__ == "__main__":
    main()
