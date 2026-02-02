import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

_CONFIGURED = False


def setup_logging(
    log_dir: Path,
    log_file: str = "rentan.log",
    level: int = logging.INFO,
) -> None:
    global _CONFIGURED
    if _CONFIGURED:
        return

    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / log_file

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    root = logging.getLogger()
    root.setLevel(level)
    root.handlers.clear()

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(level)
    console.setFormatter(formatter)

    file_handler = RotatingFileHandler(
        log_path,
        maxBytes=10_000_000,
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)

    root.addHandler(console)
    root.addHandler(file_handler)

    logging.captureWarnings(True)
    _CONFIGURED = True


class _StreamToLogger:
    def __init__(self, logger: logging.Logger, level: int):
        self.logger = logger
        self.level = level

    def write(self, message: str) -> None:
        msg = (message or "").strip()
        if msg:
            self.logger.log(self.level, msg)

    def flush(self) -> None:
        return


def setup_stdout_stderr_to_logging(logger_name: str = "STDOUT") -> None:
    """
    Redirige stdout/stderr a logging.

    - stdout -> INFO
    - stderr -> ERROR

    logger_name:
      Nombre del logger al que se envían los prints. Útil si quieres diferenciar
      "console", "ui", etc.
    """
    out_logger = logging.getLogger(logger_name)

    sys.stdout = _StreamToLogger(out_logger, logging.INFO)
    sys.stderr = _StreamToLogger(out_logger, logging.ERROR)


def install_global_excepthook(logger_name: str = "UNCAUGHT") -> None:
    """
    Captura excepciones no manejadas y las registra en logging.
    """

    def handle_exception(exc_type, exc, tb):
        logging.getLogger(logger_name).exception(
            "Excepción no capturada",
            exc_info=(exc_type, exc, tb),
        )

    sys.excepthook = handle_exception
