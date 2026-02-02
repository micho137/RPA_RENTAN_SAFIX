from __future__ import annotations

from pathlib import Path
import hashlib, re

from .client import OutlookClient
from .models import MailSummary, SaveResult


def _clean_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip()


def _digest_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


class OutlookService:
    def __init__(self, client: OutlookClient, logger):
        self.c = client
        self.log = logger

    def get_folder(self, account_display: str, folder_path: list[str]):
        store = self.c.find_store_by_display(account_display)
        folder = self.c.get_folder(store, folder_path)
        self.log.info("Using store: %s | folder: %s", store.DisplayName, "/".join(folder_path))
        return folder, store

    def list_messages(self, folder, days_back=0, only_unread=False, limit=200):
        out: list[MailSummary] = []
        for idx, item in enumerate(self.c.iter_items(folder, days_back, only_unread, limit=limit), start=1):
            out.append(MailSummary(
                entry_id=getattr(item, "EntryID", ""),
                subject=getattr(item, "Subject", "") or "",
                sender=getattr(getattr(item, "Sender", None), "Name", None),
                received=getattr(item, "ReceivedTime", None),
                has_attachments=(getattr(getattr(item, "Attachments", None), "Count", 0) or 0) > 0
            ))
        self.log.info("Messages listed: %s", len(out))
        return out

    def save_attachments(
        self,
        folder,
        out_dir: Path,
        days_back=0,
        only_unread=False,
        mark_as_read=True,
        move_to=None,
        limit: int = 5000,
    ) -> SaveResult:
        out_dir.mkdir(parents=True, exist_ok=True)
        res = SaveResult(out_dir=out_dir)

        items_iter = self.c.iter_items(folder, days_back, only_unread, limit=limit)

        for item in items_iter:
            subj = (getattr(item, "Subject", "") or "")[:80]

            try:
                self.log.info("Processing: %r", subj)

                atts = self.c.attachments(item)
                saved_now = 0

                # Subcarpeta por fecha de recepción
                received = getattr(item, "ReceivedTime", None)
                sub = out_dir / (received.strftime("%Y-%m-%d") if received else "")
                sub.mkdir(parents=True, exist_ok=True)

                for att in atts:
                    try:
                        fname = _clean_filename(getattr(att, "FileName", "") or "attachment")
                        tmp = sub / (fname + ".downloading")
                        att.SaveAsFile(str(tmp))
                        data = tmp.read_bytes()
                        h = _digest_bytes(data)
                        final = sub / f"{Path(fname).stem}__{h}{Path(fname).suffix}"
                        if final.exists():
                            tmp.unlink(missing_ok=True)
                            continue
                        tmp.rename(final)
                        saved_now += 1
                        res.attachments_saved += 1
                    except Exception as ex_att:
                        res.errors += 1
                        self.log.exception("[DL][ATT][ERROR] subj=%r err=%s", subj, ex_att)

                # marcar leído DESPUÉS de guardar
                if mark_as_read:
                    ok = self.c.mark_as_read(item)
                    if not ok:
                        res.errors += 1
                        self.log.warning("[DL][WARN] could not mark as read subj=%r", subj)

                # mover AL FINAL
                if move_to is not None:
                    ok = self.c.move_to(item, move_to)
                    if not ok:
                        res.errors += 1
                        self.log.warning("[DL][WARN] could not move subj=%r", subj)

                res.processed += 1
                self.log.info("[DL] Done subj=%r | saved_now=%s", subj, saved_now)

            except Exception as ex:
                res.errors += 1
                res.processed += 1
                self.log.exception("[DL][MAIL][ERROR] subj=%r err=%s", subj, ex)
                continue

        self.log.info("Processed=%s | Saved=%s | Errors=%s", res.processed, res.attachments_saved, res.errors)
        return res
