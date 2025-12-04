import pythoncom
from contextlib import contextmanager

@contextmanager
def com_initialized():
    """Garantiza CoInitialize/CoUninitialize para cualquier hilo que use COM."""
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()
