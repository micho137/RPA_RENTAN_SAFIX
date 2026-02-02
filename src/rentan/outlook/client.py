from __future__ import annotations

from datetime import datetime, timedelta
from typing import Iterable, Iterator, Tuple, Optional

import win32com.client
from dateutil.tz import tzlocal


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
    # Iteración estable (SIN Restrict)
    # ------------------------------
    def _build_items_snapshot(
        self,
        folder,
        days_back: int = 0,
        only_unread: bool = False,
        hard_limit: int = 5000,
    ) -> list[Tuple[str, str]]:
        """
        Devuelve lista estable de (EntryID, StoreID) en el orden actual.

        Importante:
        - Evitamos Restrict() porque es frágil (locale/format) y puede devolver 0 items
          incluso cuando sí hay correos (caso típico: solo 1 correo en el rango).
        - En su lugar, iteramos GetFirst/GetNext y filtramos en Python.
        """

        items = folder.Items
        try:
            items.Sort("[ReceivedTime]", True)  # desc
        except Exception:
            # si falla el sort, igual intentamos iterar
            pass

        since_dt: Optional[datetime] = None
        if days_back and days_back > 0:
            since_dt = datetime.now(tzlocal()) - timedelta(days=days_back)

        snap: list[Tuple[str, str]] = []

        # Iteración COM segura
        try:
            it = items.GetFirst()
        except Exception:
            it = None

        n = 0
        while it is not None:
            n += 1
            if n > hard_limit:
                break

            try:
                # Filtro unread
                if only_unread:
                    unread = bool(getattr(it, "UnRead", False))
                    if not unread:
                        it = items.GetNext()
                        continue

                # Filtro fecha
                if since_dt is not None:
                    received = getattr(it, "ReceivedTime", None)
                    if received is None:
                        it = items.GetNext()
                        continue

                    # Outlook suele retornar datetime naive; lo tratamos como local.
                    if received.tzinfo is None:
                        received_local = received.replace(tzinfo=tzlocal())
                    else:
                        received_local = received

                    if received_local < since_dt:
                        # Como está ordenado DESC, al encontrar uno más viejo,
                        # ya podemos cortar (optimiza mucho).
                        break

                entry_id = str(getattr(it, "EntryID", "") or "")
                store_id = str(getattr(it, "StoreID", "") or "")
                if entry_id and store_id:
                    snap.append((entry_id, store_id))

            except Exception:
                # saltar items dañados/no MailItem
                pass

            try:
                it = items.GetNext()
            except Exception:
                it = None

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
        return [atts.Item(i) for i in range(1, atts.Count + 1)]
