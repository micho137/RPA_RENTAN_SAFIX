from __future__ import annotations

import win32com.client
from datetime import datetime, timedelta
from typing import Iterable, Iterator, Tuple, Optional


class OutlookClient:
    """
    Cliente Outlook robusto:
    - No depende de Items.Restrict() (evita problemas regionales).
    - Filtra en Python por UnRead y ReceivedTime.
    - Itera con snapshot (EntryID, StoreID) para no romperse al marcar/mover.
    """

    MAILITEM_CLASS = 43  # Outlook MailItem

    def __init__(self):
        self._app = win32com.client.Dispatch("Outlook.Application")
        self._ns = self._app.GetNamespace("MAPI")

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

    def _build_items_snapshot(
        self,
        folder,
        days_back: int = 0,
        only_unread: bool = False,
        limit: Optional[int] = None,
    ) -> list[Tuple[str, str]]:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        snap: list[Tuple[str, str]] = []
        count = int(getattr(items, "Count", 0) or 0)

        since_delta = timedelta(days=int(days_back)) if days_back and days_back > 0 else None

        for i in range(1, count + 1):
            if limit is not None and len(snap) >= limit:
                break

            try:
                it = items.Item(i)

                # Solo MailItem
                it_class = int(getattr(it, "Class", 0) or 0)
                if it_class != self.MAILITEM_CLASS:
                    continue

                # Filtro unread
                if only_unread and not bool(getattr(it, "UnRead", False)):
                    continue

                # Filtro fecha (robusto tz-aware/naive)
                if since_delta is not None:
                    received = getattr(it, "ReceivedTime", None)
                    if received is None:
                        continue

                    # Si received es tz-aware, usamos now() con ese tz.
                    # Si es naive, usamos now() naive.
                    try:
                        if getattr(received, "tzinfo", None) is not None:
                            since_dt = datetime.now(received.tzinfo) - since_delta
                        else:
                            since_dt = datetime.now() - since_delta

                        if received < since_dt:
                            continue
                    except Exception:
                        # Si algo raro pasa con ReceivedTime, mejor no lo contamos
                        continue

                entry_id = str(getattr(it, "EntryID", "") or "")
                store_id = str(getattr(it, "StoreID", "") or "")
                if entry_id and store_id:
                    snap.append((entry_id, store_id))

            except Exception:
                continue

        return snap

    def iter_items(
        self,
        folder,
        days_back: int = 0,
        only_unread: bool = False,
        limit: Optional[int] = None,
    ) -> Iterator:
        snap = self._build_items_snapshot(folder, days_back=days_back, only_unread=only_unread, limit=limit)
        for entry_id, store_id in snap:
            try:
                yield self._ns.GetItemFromID(entry_id, store_id)
            except Exception:
                continue

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
        return [atts.Item(i) for i in range(1, atts.Count + 1)]
