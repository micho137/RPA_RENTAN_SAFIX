from __future__ import annotations

import logging
import sys
from pathlib import Path
from types import SimpleNamespace

# Stubs para importar módulos de Windows/GUI durante tests unitarios.
sys.modules.setdefault("pythoncom", SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None))
sys.modules.setdefault(
    "pyautogui",
    SimpleNamespace(
        FAILSAFE=True,
        PAUSE=0,
        write=lambda *a, **k: None,
        hotkey=lambda *a, **k: None,
        press=lambda *a, **k: None,
        keyDown=lambda *a, **k: None,
        keyUp=lambda *a, **k: None,
        locateCenterOnScreen=lambda *a, **k: None,
        click=lambda *a, **k: None,
        screenshot=lambda *a, **k: None,
    ),
)
sys.modules.setdefault(
    "pywinauto",
    SimpleNamespace(
        Application=object,
        Desktop=lambda *a, **k: None,
    ),
)
sys.modules.setdefault(
    "win32com",
    SimpleNamespace(client=SimpleNamespace(Dispatch=lambda *a, **k: None)),
)
sys.modules.setdefault(
    "win32com.client",
    SimpleNamespace(Dispatch=lambda *a, **k: None),
)

import src.workflows.invoice_pipeline as invoice_pipeline


class _DummyLogger:
    def info(self, *args, **kwargs):
        return None

    def warning(self, *args, **kwargs):
        return None


class _FakeExtractor:
    def __init__(self, *args, **kwargs):
        pass

    def process_all(self, index_csv: Path | None = None):
        if index_csv is not None:
            index_csv.parent.mkdir(parents=True, exist_ok=True)
            index_csv.write_text("zip,source,json,mode,document_id,serie\n", encoding="utf-8")
        return SimpleNamespace(zips=0, docs_json=0, errors=0)


class _CaptureTracker:
    last_instance: "_CaptureTracker | None" = None

    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)
        self.downloads: list[str] = []
        self.logs_dir = self.output_dir / "logs"
        self.logs_dir.mkdir(parents=True, exist_ok=True)
        _CaptureTracker.last_instance = self

    def add_download(self, *, original_name: str, saved_path: Path, size_bytes=None, status="OK", error=""):
        self.downloads.append(original_name)

    def save(self):
        return {
            "descargados": self.logs_dir / "descargados.xlsx",
            "procesadas": self.logs_dir / "procesadas.xlsx",
            "placas_no_encontradas": self.logs_dir / "placas_no_encontradas.xlsx",
            "doble_cc": self.logs_dir / "doble_cc.xlsx",
        }


def _patch_pipeline_dependencies(monkeypatch, tmp_path: Path):
    fake_settings = SimpleNamespace(
        log_dir=tmp_path / "logs",
        download_dir=tmp_path / "downloads",
        enable_cleanup=False,
    )
    fake_settings.log_dir.mkdir(parents=True, exist_ok=True)
    fake_settings.download_dir.mkdir(parents=True, exist_ok=True)

    monkeypatch.setattr(invoice_pipeline, "settings", fake_settings)
    monkeypatch.setattr(invoice_pipeline, "setup_logger", lambda *a, **k: _DummyLogger())
    monkeypatch.setattr(invoice_pipeline, "run_download", lambda **kwargs: SimpleNamespace(processed=0, attachments_saved=0))
    monkeypatch.setattr(invoice_pipeline, "ZipInvoiceExtractor", _FakeExtractor)
    monkeypatch.setattr(invoice_pipeline, "RunTracker", _CaptureTracker)
    return fake_settings


def test_pipeline_cleanup_disabled_returns_zero_counts(monkeypatch, tmp_path):
    _patch_pipeline_dependencies(monkeypatch, tmp_path)

    result = invoice_pipeline.run_pipeline(
        output_dir=tmp_path / "out",
        aggregate_by_id=False,
        run_safix=False,
    )

    assert result["cleanup"]["deleted_files"] == 0
    assert result["cleanup"]["deleted_dirs"] == 0


def test_pipeline_registers_downloaded_zips_recursively(monkeypatch, tmp_path):
    settings = _patch_pipeline_dependencies(monkeypatch, tmp_path)

    root_zip = settings.download_dir / "a.zip"
    nested_zip = settings.download_dir / "2026-02-09" / "b.zip"
    nested_zip.parent.mkdir(parents=True, exist_ok=True)
    root_zip.write_bytes(b"zip-a")
    nested_zip.write_bytes(b"zip-b")

    invoice_pipeline.run_pipeline(
        output_dir=tmp_path / "out",
        aggregate_by_id=False,
        run_safix=False,
    )

    tracker = _CaptureTracker.last_instance
    assert tracker is not None
    assert sorted(tracker.downloads) == ["a.zip", "b.zip"]
