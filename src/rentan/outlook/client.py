from __future__ import annotations

import win32com.client
from datetime import datetime, timedelta
from dateutil.tz import tzlocal
from typing import Iterable, Iterator, Tuple


class OutlookClient:
    def __init__(self):
        self._app = win32com.client.Dispatch("Outlook.Application")
        self._ns = self._app.GetNamespace("MAPI")

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

    # ------------------------------
    # Iteración segura (snapshot)
    # ------------------------------
    def _build_items_snapshot(self, folder, days_back: int = 0, only_unread: bool = False) -> list[Tuple[str, str]]:
        """
        Devuelve lista estable de (EntryID, StoreID) en el orden actual.
        Esto evita "out of range" cuando se marca leído o se mueve el correo.
        """
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

        snap: list[Tuple[str, str]] = []
        count = int(getattr(items, "Count", 0) or 0)

        # OJO: colección MAPI es 1-based
        for i in range(1, count + 1):
            try:
                it = items.Item(i)
                entry_id = str(getattr(it, "EntryID", "") or "")
                store_id = str(getattr(it, "StoreID", "") or "")
                if entry_id and store_id:
                    snap.append((entry_id, store_id))
            except Exception:
                # si un item se corrompe o no es MailItem, lo saltamos
                continue

        return snap

    def iter_items(self, folder, days_back: int = 0, only_unread: bool = False) -> Iterator:
        """
        Itera MailItems vía snapshot: no se rompe si cambias UnRead o haces Move().
        """
        snap = self._build_items_snapshot(folder, days_back=days_back, only_unread=only_unread)
        for entry_id, store_id in snap:
            try:
                yield self._ns.GetItemFromID(entry_id, store_id)
            except Exception:
                continue

    # Operaciones sobre MailItem (devuelven bool)
    @staticmethod
    def mark_as_read(mail_item) -> bool:
        try:
            if getattr(mail_item, "UnRead", False):
                mail_item.UnRead = False
                mail_item.Save()
            return True
        except Exception:
            return False

    @staticmethod
    def move_to(mail_item, folder) -> bool:
        try:
            mail_item.Move(folder)
            return True
        except Exception:
            return False

    @staticmethod
    def attachments(mail_item) -> Iterable:
        atts = getattr(mail_item, "Attachments", None)
        if not atts or getattr(atts, "Count", 0) == 0:
            return []
        # Colección MAPI es 1-based
        return [atts.Item(i) for i in range(1, atts.Count + 1)]