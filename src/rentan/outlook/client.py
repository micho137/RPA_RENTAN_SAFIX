from __future__ import annotations

import win32com.client
from datetime import datetime, timedelta
from typing import Iterable, Iterator, Tuple, Optional


class OutlookClient:
    """
    Cliente Outlook robusto:
    - NO usa Items.Restrict() para evitar problemas de formato regional y propiedades MAPI.
    - Aplica filtros (days_back / only_unread) en Python.
    - Itera con snapshot (EntryID, StoreID) para no romperse si marcas leído o mueves correos.
    """

    MAILITEM_CLASS = 43  # Outlook MailItem

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
    def _build_items_snapshot(
        self,
        folder,
        days_back: int = 0,
        only_unread: bool = False,
        limit: Optional[int] = None,
    ) -> list[Tuple[str, str]]:
        """
        Devuelve lista estable de (EntryID, StoreID) en el orden actual.
        Evita "out of range" cuando se marca leído o se mueve el correo.

        Aplica filtros en Python:
        - days_back: filtra por ReceivedTime >= now - days_back
        - only_unread: filtra por item.UnRead == True

        limit: corta el snapshot al número máximo de items.
        """
        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        snap: list[Tuple[str, str]] = []
        count = int(getattr(items, "Count", 0) or 0)

        since_dt = None
        if days_back and days_back > 0:
            since_dt = datetime.now() - timedelta(days=int(days_back))

        # OJO: colección MAPI es 1-based
        for i in range(1, count + 1):
            if limit is not None and len(snap) >= limit:
                break

            try:
                it = items.Item(i)

                # Asegura MailItem
                it_class = int(getattr(it, "Class", 0) or 0)
                if it_class != self.MAILITEM_CLASS:
                    continue

                # Filtro unread (en Python)
                if only_unread:
                    if not bool(getattr(it, "UnRead", False)):
                        continue

                # Filtro fecha (en Python)
                if since_dt is not None:
                    received = getattr(it, "ReceivedTime", None)
                    if received is None:
                        continue

                    # ReceivedTime suele ser datetime “naive” desde COM
                    # Comparamos naive vs naive (datetime.now() es naive)
                    try:
                        if received < since_dt:
                            continue
                    except Exception:
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
        """
        Itera MailItems vía snapshot: no se rompe si cambias UnRead o haces Move().
        """
        snap = self._build_items_snapshot(folder, days_back=days_back, only_unread=only_unread, limit=limit)
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
