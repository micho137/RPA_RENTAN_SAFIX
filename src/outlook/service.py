from __future__ import annotations
from pathlib import Path
from datetime import date, datetime
import hashlib, re
import json
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
        self.log.info(f"Using store: {store.DisplayName} | folder: {'/'.join(folder_path)}")
        return folder, store

    def list_messages(
        self,
        folder,
        days_back=0,
        only_unread=False,
        limit=200,
        date_from: date | None = None,
        date_to: date | None = None,
    ):
        out: list[MailSummary] = []
        for idx, item in enumerate(
            self.c.iter_items(
                folder,
                days_back=days_back,
                only_unread=only_unread,
                date_from=date_from,
                date_to=date_to,
            ),
            start=1,
        ):
            if idx > limit: break
            out.append(MailSummary(
                entry_id=getattr(item, "EntryID", ""),
                subject=getattr(item, "Subject", "") or "",
                sender=getattr(getattr(item, "Sender", None), "Name", None),
                received=getattr(item, "ReceivedTime", None),
                has_attachments=(getattr(getattr(item, "Attachments", None), "Count", 0) or 0) > 0
            ))
        self.log.info(f"Messages listed: {len(out)}")
        return out

    def save_attachments(
        self,
        folder,
        out_dir: Path,
        days_back=0,
        only_unread=False,
        mark_as_read=True,
        move_to=None,
        date_from: date | None = None,
        date_to: date | None = None,
    ) -> SaveResult:
        out_dir.mkdir(parents=True, exist_ok=True)
        res = SaveResult(out_dir=out_dir)
        manifest_path = out_dir / "_manifest.json"
        manifest: dict[str, str] = {}
        if manifest_path.exists():
            try:
                manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            except Exception:
                manifest = {}

        for item in self.c.iter_items(
            folder,
            days_back=days_back,
            only_unread=only_unread,
            date_from=date_from,
            date_to=date_to,
        ):
            subj = (getattr(item, "Subject", "") or "")[:80]
            self.log.info(f"Processing: {subj!r}")

            atts = self.c.attachments(item)
            saved_now = 0
            entry_id = getattr(item, "EntryID", "") or ""

            # Subcarpeta por fecha de recepción
            received = getattr(item, "ReceivedTime", None)
            sub = out_dir / (datetime.fromtimestamp(int(received.timestamp())).strftime("%Y-%m-%d") if received else "")
            sub.mkdir(parents=True, exist_ok=True)

            for att in atts:
                fname = _clean_filename(att.FileName or "attachment")
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
                if entry_id:
                    manifest[final.stem] = entry_id

            if mark_as_read:
                self.c.mark_as_read(item)
            else:
                # Acceder a adjuntos via COM puede causar que Outlook
                # auto-marque el correo como leído. Revertimos ese cambio
                # para que no-leídos sigan disponibles en corridas futuras.
                try:
                    if not getattr(item, "UnRead", True):
                        item.UnRead = True
                        item.Save()
                except Exception:
                    pass
            if move_to is not None:
                self.c.move_to(item, move_to)

            res.processed += 1

        try:
            manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
            res.manifest_path = manifest_path
        except Exception:
            pass

        self.log.info(f"Processed={res.processed} | Saved={res.attachments_saved}")
        return res
