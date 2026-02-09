from __future__ import annotations

from pathlib import Path

from src.core.progress_store import ProgressStore


def test_progress_store_persists_downloaded_and_processed(tmp_path):
    logs_dir = tmp_path / "logs"
    store = ProgressStore(logs_dir)

    store.register_downloaded(document_id="DEFL100", placa="ABC123", source_key="k1")
    store.mark_processed(document_id="DEFL100", placa="ABC123", status="OK", source_key="k1")

    store2 = ProgressStore(logs_dir)
    processed = store2.get_processed_ok_ids()

    assert "DEFL100" in processed
    assert (logs_dir / "resume_state.json").exists()
    assert (logs_dir / "facturas_descargadas.xlsx").exists()
    assert (logs_dir / "facturas_procesadas_global.xlsx").exists()


def test_progress_store_only_ok_is_resumable(tmp_path):
    logs_dir = tmp_path / "logs"
    store = ProgressStore(logs_dir)

    store.register_downloaded(document_id="DEFL101", placa="AAA111", source_key="k1")
    store.mark_processed(document_id="DEFL101", placa="AAA111", status="FAIL", error="boom", source_key="k1")
    store.register_downloaded(document_id="DEFL102", placa="BBB222", source_key="k2")
    store.mark_processed(document_id="DEFL102", placa="BBB222", status="OK", source_key="k2")

    processed = store.get_processed_ok_ids()
    assert "DEFL102" in processed
    assert "DEFL101" not in processed
