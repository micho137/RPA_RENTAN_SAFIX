# src/workflows/download_attachments.py
from pathlib import Path
from datetime import date
from src.core.logging_config import setup_logger
from src.core.com_init import com_initialized
from src.config import settings
from src.outlook.client import OutlookClient
from src.outlook.service import OutlookService


def run_download(
    *,
    days_back: int,
    only_unread: bool,
    mark_as_read: bool,
    move_to_processed: bool,
    date_from: date | None = None,
    date_to: date | None = None,
):
    logger = setup_logger("outlook_bot", settings.log_dir)

    logger.info(
        "[DL][FLAGS] only_unread=%s days_back=%s date_from=%s date_to=%s mark_as_read=%s move_to_processed=%s",
        only_unread, days_back, date_from, date_to, mark_as_read, move_to_processed
    )

    with com_initialized():
        client = OutlookClient()
        service = OutlookService(client, logger)

        folder, store = service.get_folder(
            account_display=settings.outlook_account,
            folder_path=[settings.source_folder]
        )

        move_to = None
        if move_to_processed:
            move_to = service.c.get_folder(store, [settings.processed_folder])

        res = service.save_attachments(
            folder=folder,
            out_dir=settings.download_dir,
            days_back=days_back,
            only_unread=only_unread,
            mark_as_read=mark_as_read,
            move_to=move_to,
            date_from=date_from,
            date_to=date_to,
        )

    return res
