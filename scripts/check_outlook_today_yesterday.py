from __future__ import annotations

import argparse
from datetime import date, datetime, timedelta

from src.config import settings
from src.outlook.client import OutlookClient


def _parse_date(raw: str) -> date:
    return datetime.strptime(raw, "%Y-%m-%d").date()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Lista correos de Outlook por dias atras o por rango de fechas."
    )
    parser.add_argument(
        "--days-back",
        type=int,
        default=1,
        help="Dias hacia atras (incluye hoy). Ignorado si usas --from y --to.",
    )
    parser.add_argument("--from", dest="from_date", type=str, help="Fecha inicial YYYY-MM-DD")
    parser.add_argument("--to", dest="to_date", type=str, help="Fecha final YYYY-MM-DD")
    args = parser.parse_args()

    account = (settings.outlook_account or "").strip()
    folder_path_raw = (settings.source_folder or "").strip()
    if not account:
        print("[ERROR] OUTLOOK_ACCOUNT no esta configurado en .env")
        return 1
    if not folder_path_raw:
        print("[ERROR] OUTLOOK_FOLDER no esta configurado en .env")
        return 1

    use_range = bool(args.from_date and args.to_date)
    if (args.from_date and not args.to_date) or (args.to_date and not args.from_date):
        print("[ERROR] Debes enviar ambas fechas: --from y --to")
        return 1

    if use_range:
        try:
            range_start = _parse_date(args.from_date)
            range_end = _parse_date(args.to_date)
        except ValueError:
            print("[ERROR] Formato invalido. Usa YYYY-MM-DD")
            return 1
        if range_end < range_start:
            print("[ERROR] --to no puede ser menor que --from")
            return 1
        query_days_back = max(1, (date.today() - range_start).days + 1)
    else:
        query_days_back = max(0, int(args.days_back))
        range_end = date.today()
        range_start = range_end - timedelta(days=query_days_back)

    folder_path = [p.strip() for p in folder_path_raw.split("/") if p.strip()]

    print(f"[INFO] Cuenta: {account}")
    print(f"[INFO] Carpeta: {'/'.join(folder_path)}")
    if use_range:
        print(f"[INFO] Rango solicitado: {range_start} -> {range_end}")
    else:
        print(f"[INFO] days_back={query_days_back} (hoy incluido)")

    client = OutlookClient()
    store = client.find_store_by_display(account)
    folder = client.get_folder(store, folder_path)

    items = list(client.iter_items(folder, days_back=query_days_back, only_unread=False))

    by_day: dict[date, list[str]] = {}
    for item in items:
        received = getattr(item, "ReceivedTime", None)
        subject = (getattr(item, "Subject", "") or "").strip() or "(sin asunto)"
        if received is None:
            continue

        item_day = received.date()
        if item_day < range_start or item_day > range_end:
            continue

        line = f"- {received:%Y-%m-%d %H:%M} | {subject}"
        by_day.setdefault(item_day, []).append(line)

    print("=== CORREOS EN RANGO ===")
    cursor = range_end
    total_printed = 0
    while cursor >= range_start:
        rows = by_day.get(cursor, [])
        print(f"[{cursor}] total={len(rows)}")
        if rows:
            for row in rows:
                print(row)
                total_printed += 1
        else:
            print("(sin correos)")
        cursor = cursor - timedelta(days=1)

    print(f"[INFO] Totales -> impresos={total_printed} | "f"iterados(days_back={query_days_back})={len(items)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
