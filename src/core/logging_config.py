import logging
import sys
from datetime import datetime
from pathlib import Path
from . import com_init  # solo para mantener orden en el paquete

_LOG_INITIALIZED = False
_ORIG_STDOUT = sys.stdout
_ORIG_STDERR = sys.stderr


def _is_valid_stream(stream) -> bool:
    return stream is not None and hasattr(stream, "write")


class _StreamToLogger:
    def __init__(self, logger: logging.Logger, level: int):
        self._logger = logger
        self._level = level
        self._in_write = False

    def write(self, message: str) -> None:
        msg = (message or "").rstrip()
        if not msg or self._in_write:
            return
        try:
            self._in_write = True
            self._logger.log(self._level, msg)
        except Exception:
            # Fallback duro para evitar recursion de logging.
            if _is_valid_stream(_ORIG_STDERR):
                try:
                    _ORIG_STDERR.write(msg + "\n")
                except Exception:
                    pass
        finally:
            self._in_write = False

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

    console = None
    if _is_valid_stream(_ORIG_STDOUT):
        console = logging.StreamHandler(_ORIG_STDOUT)
        console.setFormatter(fmt)

    date_tag = datetime.now().strftime("%Y-%m-%d")
    file_handler = logging.FileHandler(
        log_dir / f"app_{date_tag}.log",
        encoding="utf-8"
    )
    file_handler.setFormatter(fmt)

    root.setLevel(level)
    if console is not None:
        root.addHandler(console)
    root.addHandler(file_handler)
    root.propagate = False

    _LOG_INITIALIZED = True
    return root


def redirect_std_streams(logger: logging.Logger, include_stderr: bool = False) -> None:
    sys.stdout = _StreamToLogger(logger, logging.INFO)
    # No redirigir stderr por defecto para evitar recursion con logging.handleError.
    if include_stderr:
        sys.stderr = _StreamToLogger(logger, logging.ERROR)


def setup_logger(name: str, log_dir: Path, level=logging.INFO) -> logging.Logger:
    init_logging(log_dir, level=level)
    logger = logging.getLogger(name)
    logger.setLevel(level)
    return logger
