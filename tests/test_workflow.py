import types
from pathlib import Path
from src.outlook.service import OutlookService
from src.outlook.client import OutlookClient

class FakeAttachment:
    """Simula un archivo adjunto de correo"""
    def __init__(self, name, data=b"hello"):
        self.FileName = name
        self._data = data

    def SaveAsFile(self, p):
        print(f"   Guardando adjunto simulado: {self.FileName} -> {p}")
        Path(p).write_bytes(self._data)


class FakeMail:
    """Simula un correo electrónico de Outlook"""
    def __init__(self, subject="Test", unread=True, with_atts=True):
        self.Subject = subject
        self.UnRead = unread
        self._atts = [FakeAttachment("a.txt"), FakeAttachment("b.pdf")] if with_atts else []

    @property
    def Attachments(self):
        """Simula la colección de adjuntos COM"""
        class A:
            def __init__(self, ls): self.ls = ls
            @property
            def Count(self): return len(self.ls)
            def Item(self, i): return self.ls[i - 1]
        return A(self._atts)


class FakeClient(OutlookClient):
    """Cliente falso para pruebas sin Outlook real"""
    def __init__(self): pass

    def iter_items(self, folder, days_back=0, only_unread=False):
        print(f"[FakeClient] Iterando correos simulados (days_back={days_back}, only_unread={only_unread})")
        yield FakeMail("uno")              # con adjuntos
        yield FakeMail("dos", with_atts=False)  # sin adjuntos

    @staticmethod
    def mark_as_read(mail_item):
        mail_item.UnRead = False
        print(f"   Marcando como leído: {mail_item.Subject}")

    @staticmethod
    def move_to(mail_item, folder):
        print(f"   Moviendo correo (simulado): {mail_item.Subject} -> {folder}")


def test_service_saves_attachments(tmp_path):
    print("\n[TEST] Ejecutando prueba completa de guardado de adjuntos...")
    logger = types.SimpleNamespace(info=lambda *a, **k: print("[LOG]", *a))
    svc = OutlookService(FakeClient(), logger)

    print("Creando carpeta temporal de destino:", tmp_path)
    res = svc.save_attachments(folder=object(), out_dir=tmp_path, mark_as_read=True)

    print(f"Correos procesados: {res.processed}")
    print(f"Adjuntos guardados: {res.attachments_saved}")

    files = list(tmp_path.rglob("*"))
    print("Archivos generados:")
    for f in files:
        print(" -", f.name)

    assert res.processed == 2
    assert res.attachments_saved == 2
    assert any(p.suffix == ".txt" for p in files)
    assert any(p.suffix == ".pdf" for p in files)

    print("[OK] Flujo de guardado de adjuntos completado correctamente ✅")
