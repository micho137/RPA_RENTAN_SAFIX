from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

@dataclass
class MailSummary:
    entry_id: str
    subject: str
    sender: str | None
    received: datetime | None
    has_attachments: bool

@dataclass
class SaveResult:
    processed: int = 0
    attachments_saved: int = 0
    out_dir: Path | None = None
