from __future__ import annotations

import json
import sys
from pathlib import Path
from types import SimpleNamespace

# Stubs para importar módulos Windows/GUI en tests.
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

from src.workflows.safix_automation import (
    _manifest_entry_ids_for_invoice,
    load_attachment_manifest,
    normalize_placa,
)


def test_normalize_placa_removes_separators() -> None:
    assert normalize_placa(" abc-123 ") == "ABC123"
    assert normalize_placa("ab c.1_2/3") == "ABC123"


def test_load_attachment_manifest_supports_legacy_and_list(tmp_path: Path) -> None:
    manifest_path = tmp_path / "_manifest.json"
    manifest_path.write_text(
        json.dumps(
            {
                "zip_a": "ENTRY_A",
                "zip_b": ["ENTRY_B1", "ENTRY_B2", "ENTRY_B1", ""],
                "zip_c": 12345,
                "": "EMPTY_KEY_IGNORED",
            }
        ),
        encoding="utf-8",
    )

    out = load_attachment_manifest(tmp_path)
    assert out["zip_a"] == ["ENTRY_A"]
    assert out["zip_b"] == ["ENTRY_B1", "ENTRY_B2"]
    assert "zip_c" not in out
    assert "" not in out


def test_manifest_entry_ids_for_invoice_resolves_from_source_path(tmp_path: Path) -> None:
    extract_root = tmp_path / "extract"
    source = extract_root / "zip_20260316__abcd1234" / "inner" / "invoice.xml"
    source.parent.mkdir(parents=True, exist_ok=True)
    source.write_text("<x/>", encoding="utf-8")

    invoice = {"_source_path": str(source)}
    manifest = {"zip_20260316__abcd1234": ["ENTRY_1", "ENTRY_2"]}

    ids = _manifest_entry_ids_for_invoice(invoice, manifest, extract_root)
    assert ids == ["ENTRY_1", "ENTRY_2"]

