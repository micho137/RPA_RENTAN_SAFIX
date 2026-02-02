from __future__ import annotations

import win32com.client
from datetime import datetime, timedelta
from dateutil.tz import tzlocal
from typing import Iterable, Iterator, Tuple, List, Optional


MAILITEM_CLASS = 43  # Outlook.MailItem


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
    def _build_items_snapshot(
        self,
        folder,
        days_back: int = 0,
        only_unread: bool = False,
        limit: int = 10000,
    ) -> List[Tuple[str, str]]:
        """
        Devuelve lista estable de (EntryID, StoreID) en el orden actual.

        Cambio clave:
        - NO usamos items.Item(i) porque Outlook COM falla de forma intermitente,
          especialmente con colecciones pequeñas (1 item) o Restrict().
        - Usamos GetFirst()/GetNext() que es la forma robusta de iterar.
        """
        items = folder.Items
        # Ordenar por fecha descendente (más reciente primero)
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass

        # Aplicar filtros si corresponde
        if days_back > 0 or only_unread:
            items = self._restrict_items(items, days_back=days_back, only_unread=only_unread)

        snap: List[Tuple[str, str]] = []

        try:
            it = items.GetFirst()
        except Exception:
            it = None

        n = 0
        while it is not None and n < limit:
            n += 1
            try:
                # Filtrar solo MailItem real
                if int(getattr(it, "Class", 0) or 0) != MAILITEM_CLASS:
                    try:
                        it = items.GetNext()
                    except Exception:
                        break
                    continue

                entry_id = str(getattr(it, "EntryID", "") or "")
                store_id = str(getattr(it, "StoreID", "") or "")
                if entry_id and store_id:
                    snap.append((entry_id, store_id))

                try:
                    it = items.GetNext()
                except Exception:
                    break

            except Exception:
                # Si este item falla, intentamos avanzar
                try:
                    it = items.GetNext()
                except Exception:
                    break

        return snap

    def _restrict_items(self, items, days_back: int, only_unread: bool):
        """
        Restrict robusto usando @SQL (DASL).
        Esto evita problemas de formato regional con ReceivedTime.
        """
        clauses = []

        if days_back > 0:
            since = (datetime.now(tzlocal()) - timedelta(days=days_back))
            # Formato recomendado para DASL: YYYY-MM-DD HH:MM
            since_str = since.strftime("%Y-%m-%d %H:%M")
            # datereceived
            clauses.append(f"\"urn:schemas:httpmail:datereceived\" >= '{since_str}'")

        if only_unread:
            # read = 0 => no leído
            clauses.append("\"urn:schemas:httpmail:read\" = 0")

        if not clauses:
            return items

        sql = "@SQL=" + " AND ".join(clauses)

        try:
            return items.Restrict(sql)
        except Exception:
            # Fallback a Restrict clásico (menos robusto)
            filters = []
            if days_back > 0:
                since = (datetime.now(tzlocal()) - timedelta(days=days_back))
                since_str = since.strftime("%m/%d/%Y %I:%M %p")
                filters.append(f"[ReceivedTime] >= '{since_str}'")
            if only_unread:
                filters.append("[UnRead] = True")

            if filters:
                try:
                    return items.Restrict(" AND ".join(filters))
                except Exception:
                    return items

        return items

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
