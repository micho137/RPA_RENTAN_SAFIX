from __future__ import annotations

import logging


from src.rentan.core.com_init import com_initialized
from src.rentan.config.config import settings
from src.rentan.outlook.client import OutlookClient
from src.rentan.outlook.service import OutlookService


logger = logging.getLogger(__name__)


def run_download(
    *,
    days_back: int,
    only_unread: bool,
    mark_as_read: bool,
    move_to_processed: bool,
):
    """
    Descarga adjuntos desde Outlook según flags.
    Todo el logging va al logger unificado (configurado en el entrypoint).
    """
    logger.info(
        "[DL][FLAGS] only_unread=%s days_back=%s mark_as_read=%s move_to_processed=%s",
        only_unread,
        days_back,
        mark_as_read,
        move_to_processed,
    )

    with com_initialized():
        client = OutlookClient()
        service = OutlookService(client, logger)

        folder, store = service.get_folder(
            account_display=settings.outlook_account,
            folder_path=[settings.source_folder],
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
        )

    logger.info(
        "[DL] Finished | processed=%s attachments_saved=%s errors=%s out_dir=%s",
        res.processed,
        res.attachments_saved,
        res.errors,
        res.out_dir,
    )
    return res
