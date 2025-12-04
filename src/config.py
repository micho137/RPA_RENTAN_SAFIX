from dataclasses import dataclass
import os
from dotenv import load_dotenv
from pathlib import Path

load_dotenv()

@dataclass(frozen=True)
class Settings:
    outlook_account: str = os.getenv("OUTLOOK_ACCOUNT")
    source_folder: str = os.getenv("OUTLOOK_FOLDER")
    download_dir: Path = Path(os.getenv("DOWNLOAD_DIR", "./downloads")).resolve()
    only_unread: bool = os.getenv("ONLY_UNREAD", "false").lower() == "true"
    days_back: int = int(os.getenv("DAYS_BACK", "30"))
    mark_as_read: bool = os.getenv("MARK_AS_READ", "true").lower() == "true"
    move_to_processed: bool = os.getenv("MOVE_TO_PROCESSED", "false").lower() == "true"
    processed_folder: str = os.getenv("PROCESSED_FOLDER", "Procesados")
    log_dir: Path = Path(os.getenv("LOG_DIR", "./logs")).resolve()

settings = Settings()
settings.download_dir.mkdir(parents=True, exist_ok=True)
settings.log_dir.mkdir(parents=True, exist_ok=True)
