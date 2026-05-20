from __future__ import annotations

import importlib


def test_doc_id_defaults_accept_defl_and_fpfl(monkeypatch) -> None:
    monkeypatch.delenv("DOC_ID_VALID_SERIES", raising=False)
    mod = importlib.import_module("src.processing.doc_id")
    mod = importlib.reload(mod)

    assert mod.parse_any_id("DEFL-12345678") == ("DEFL", "12345678", "DEFL-12345678")
    assert mod.parse_any_id("FPFL12345678") == ("FPFL", "12345678", "FPFL-12345678")
    assert mod.parse_any_id("DENC-12345678") is None


def test_doc_id_can_be_overridden_by_env(monkeypatch) -> None:
    monkeypatch.setenv("DOC_ID_VALID_SERIES", "DEFL,ABCD")
    mod = importlib.import_module("src.processing.doc_id")
    mod = importlib.reload(mod)

    assert mod.parse_any_id("ABCD-12345678") == ("ABCD", "12345678", "ABCD-12345678")
    assert mod.parse_any_id("FPFL-12345678") is None

