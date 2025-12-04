import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from . import com_init  # solo para mantener orden en el paquete

def setup_logger(name: str, log_dir: Path, level=logging.INFO) -> logging.Logger:
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger
    logger.setLevel(level)

    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    console = logging.StreamHandler()
    console.setFormatter(fmt)
    logger.addHandler(console)

    file_handler = RotatingFileHandler(
        log_dir / f"{name}.log",
        maxBytes=2_000_000,
        backupCount=5,
        encoding="utf-8"
    )
    file_handler.setFormatter(fmt)
    logger.addHandler(file_handler)

    return logger
