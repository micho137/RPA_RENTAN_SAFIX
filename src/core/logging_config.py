import logging
import sys
from datetime import datetime
from pathlib import Path
from . import com_init  # solo para mantener orden en el paquete

_LOG_INITIALIZED = False
_ORIG_STDOUT = sys.stdout
_ORIG_STDERR = sys.stderr


class _StreamToLogger:
    def __init__(self, logger: logging.Logger, level: int):
        self._logger = logger
        self._level = level

    def write(self, message: str) -> None:
        msg = (message or "").rstrip()
        if msg:
            self._logger.log(self._level, msg)

    def flush(self) -> None:
        return


def init_logging(log_dir: Path, level=logging.INFO) -> logging.Logger:
    global _LOG_INITIALIZED
    root = logging.getLogger()

    if _LOG_INITIALIZED:
        return root

    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)

    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    console = logging.StreamHandler(_ORIG_STDOUT)
    console.setFormatter(fmt)

    date_tag = datetime.now().strftime("%Y-%m-%d")
    file_handler = logging.FileHandler(
        log_dir / f"app_{date_tag}.log",
        encoding="utf-8"
    )
    file_handler.setFormatter(fmt)

    root.setLevel(level)
    root.addHandler(console)
    root.addHandler(file_handler)
    root.propagate = False

    _LOG_INITIALIZED = True
    return root


def redirect_std_streams(logger: logging.Logger) -> None:
    sys.stdout = _StreamToLogger(logger, logging.INFO)
    sys.stderr = _StreamToLogger(logger, logging.ERROR)


def setup_logger(name: str, log_dir: Path, level=logging.INFO) -> logging.Logger:
    init_logging(log_dir, level=level)
    logger = logging.getLogger(name)
    logger.setLevel(level)
    return logger
