from pathlib import Path
from src.core.logging_config import setup_logger
from src.core.com_init import com_initialized
from src.config import settings
from src.outlook.client import OutlookClient
from src.outlook.service import OutlookService

def run_download():
    logger = setup_logger("outlook_bot", settings.log_dir)

    with com_initialized():
        client = OutlookClient()
        service = OutlookService(client, logger)

        folder, store = service.get_folder(
            account_display=settings.outlook_account,
            folder_path=[settings.source_folder]
        )

        move_to = None
        if settings.move_to_processed:
            move_to = service.c.get_folder(store, [settings.processed_folder])

        res = service.save_attachments(
            folder=folder,
            out_dir=settings.download_dir,
            days_back=settings.days_back,
            only_unread=settings.only_unread,
            mark_as_read=settings.mark_as_read,
            move_to=move_to
        )

    return res
