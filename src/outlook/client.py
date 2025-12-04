from __future__ import annotations
import win32com.client
from datetime import datetime, timedelta
from dateutil.tz import tzlocal
from typing import Iterable

class OutlookClient:
    def __init__(self):
        self._ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Stores y carpetas
    def find_store_by_display(self, display: str):
        for store in self._ns.Stores:
            if display.lower() in store.DisplayName.lower():
                return store
        raise RuntimeError(f"Outlook store '{display}' not found")

    def get_folder(self, store, path: list[str]):
        folder = store.GetRootFolder()
        for p in path:
            folder = folder.Folders[p]
        return folder

    # Items + filtros
    def iter_items(self, folder, days_back: int = 0, only_unread: bool = False):
        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        filters = []
        if days_back > 0:
            since = (datetime.now(tzlocal()) - timedelta(days=days_back))
            since_str = since.strftime("%m/%d/%Y %I:%M %p")
            filters.append(f"[ReceivedTime] >= '{since_str}'")
        if only_unread:
            filters.append("[UnRead] = True")
        if filters:
            items = items.Restrict(" AND ".join(filters))

        # Colección MAPI es 1-based
        for i in range(1, items.Count + 1):
            yield items.Item(i)

    # Operaciones sobre MailItem
    @staticmethod
    def mark_as_read(mail_item):
        if getattr(mail_item, "UnRead", False):
            mail_item.UnRead = False
            mail_item.Save()

    @staticmethod
    def move_to(mail_item, folder):
        mail_item.Move(folder)

    @staticmethod
    def attachments(mail_item) -> Iterable:
        atts = getattr(mail_item, "Attachments", None)
        return [] if not atts or atts.Count == 0 else [atts.Item(i) for i in range(1, atts.Count + 1)]
