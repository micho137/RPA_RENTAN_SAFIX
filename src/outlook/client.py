from __future__ import annotations
import win32com.client
from datetime import date, datetime, timedelta
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
    def iter_items(
        self,
        folder,
        days_back: int = 0,
        only_unread: bool = False,
        date_from: date | None = None,
        date_to: date | None = None,
    ):
        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        # Mantener Restrict solo para UnRead: es estable entre locales.
        # El filtro de fechas se aplica en Python para evitar errores por formato regional.
        filters = []
        if only_unread:
            filters.append("[UnRead] = True")
        if filters:
            items = items.Restrict(" AND ".join(filters))

        # Colección MAPI es 1-based. Hacemos snapshot para evitar
        # "index out of range" cuando se marca como leído y cambia el filtro.
        snapshot = [items.Item(i) for i in range(1, items.Count + 1)]

        local_tz = tzlocal()
        if date_from is not None and date_to is not None:
            since = datetime.combine(date_from, datetime.min.time()).replace(tzinfo=local_tz)
            until_exclusive = datetime.combine(date_to + timedelta(days=1), datetime.min.time()).replace(tzinfo=local_tz)
        elif days_back > 0:
            since = (datetime.now(local_tz) - timedelta(days=days_back))
            until_exclusive = None
        else:
            since = None
            until_exclusive = None

        for item in snapshot:
            if since is not None:
                received = getattr(item, "ReceivedTime", None)
                if received is None:
                    continue
                if received.tzinfo is None:
                    received = received.replace(tzinfo=local_tz)
                if received < since:
                    continue
                if until_exclusive is not None and received >= until_exclusive:
                    continue
            yield item

    def get_item_by_id(self, entry_id: str):
        return self._ns.GetItemFromID(entry_id)

    def mark_as_read_by_id(self, entry_id: str):
        item = self.get_item_by_id(entry_id)
        if item is not None:
            self.mark_as_read(item)

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
