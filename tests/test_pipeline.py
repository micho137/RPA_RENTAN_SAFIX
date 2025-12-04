import sys
import json
import zipfile
from pathlib import Path
import pytest

# Asegura que 'src' quede importable
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.workflows.invoice_pipeline import run_pipeline
from src.processing.zip_invoice_extractor import ZipInvoiceExtractor
from src.config import settings


def _make_ubl_invoice_xml(serie: str, numero: str) -> str:
    """XML UBL mínimo con <Invoice> que el extractor sabe leer."""
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Invoice xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
         xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
         xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2">
  <cbc:ID>{serie}{numero}</cbc:ID>
  <cbc:UUID>CUFE-DEMO-123</cbc:UUID>
  <cbc:IssueDate>2025-10-15</cbc:IssueDate>
  <cbc:IssueTime>19:26:06</cbc:IssueTime>
  <cbc:DocumentCurrencyCode>COP</cbc:DocumentCurrencyCode>

  <cac:AccountingSupplierParty>
    <cac:Party>
      <cac:PartyName><cbc:Name>F2X S.A.S.</cbc:Name></cac:PartyName>
    </cac:Party>
    <cac:PartyTaxScheme><cbc:CompanyID>900219834</cbc:CompanyID></cac:PartyTaxScheme>
  </cac:AccountingSupplierParty>

  <cac:AccountingCustomerParty>
    <cac:Party>
      <cac:PartyName><cbc:Name>Renting de Antioquia EICE</cbc:Name></cac:PartyName>
      <cac:Contact><cbc:Telephone>3016904190</cbc:Telephone></cac:Contact>
      <cac:PhysicalLocation>
        <cac:Address>
          <cac:AddressLine><cbc:Line>Cra 43A NO.19127</cbc:Line></cac:AddressLine>
          <cbc:CityName>Medellín</cbc:CityName>
          <cac:Country><cbc:Name>Colombia</cbc:Name></cac:Country>
        </cac:Address>
      </cac:PhysicalLocation>
    </cac:Party>
    <cac:PartyTaxScheme><cbc:CompanyID>900285704</cbc:CompanyID></cac:PartyTaxScheme>
  </cac:AccountingCustomerParty>

  <cac:PaymentMeans>
    <cbc:PaymentDueDate>2025-10-15</cbc:PaymentDueDate>
  </cac:PaymentMeans>

  <cac:InvoiceLine>
    <cbc:ID>1</cbc:ID>
    <cbc:InvoicedQuantity unitCode="94">1</cbc:InvoicedQuantity>
    <cac:Price><cbc:PriceAmount>16100.00</cbc:PriceAmount></cac:Price>
    <cac:Item><cbc:Description>Placa: GEM560::Peaje: AMAGA::Categoria: 1::Fecha del paso: 16/10/2025 00:19:00</cbc:Description></cac:Item>
    <cbc:LineExtensionAmount>16100.00</cbc:LineExtensionAmount>
  </cac:InvoiceLine>

  <cac:LegalMonetaryTotal>
    <cbc:LineExtensionAmount>16100.00</cbc:LineExtensionAmount>
    <cbc:PayableAmount>16100.00</cbc:PayableAmount>
  </cac:LegalMonetaryTotal>
</Invoice>
""".strip()


def _make_zip_with_file(zip_path: Path, inner_relpath: str, content: bytes):
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(inner_relpath, content)


class _DummyDL:
    def __init__(self, processed=2, attachments_saved=2):
        self.processed = processed
        self.attachments_saved = attachments_saved


def test_run_pipeline_xml_and_pdf_regex(tmp_path, monkeypatch):
    """
    Cobertura del workflow completo:
      - ZIP A: trae XML UBL válido (serie DEFL) -> ruta XML-first
      - ZIP B: no trae XML, solo PDF -> se fuerza texto (serie DENC) via monkeypatch pdf_to_text()
    """
    # Estructura temporal
    downloads = tmp_path / "downloads"
    extract   = tmp_path / "unzipped"
    out_json  = tmp_path / "json"
    out_text  = tmp_path / "text"
    index_csv = out_json / "index.csv"
    for p in (downloads, extract, out_json, out_text):
        p.mkdir(parents=True, exist_ok=True)

    # ZIP A (XML UBL)
    xml = _make_ubl_invoice_xml("DEFL", "62045612").encode("utf-8")
    _make_zip_with_file(downloads / "a_xml.zip", "docs/invoice.xml", xml)

    # ZIP B (solo PDF) — el contenido no importa, usaremos monkeypatch de pdf_to_text
    _make_zip_with_file(downloads / "b_pdf.zip", "docs/factura.pdf", b"%PDF-1.4\n%EOF\n")

    # Redirige el directorio de descargas del pipeline
    monkeypatch.setattr(settings, "download_dir", downloads)

    # Evita dependencia de Outlook: parchea run_download()
    import src.workflows.download_attachments as dl
    monkeypatch.setattr(dl, "run_download", lambda: _DummyDL())

    # Fuerza texto "OCR" para el PDF con un ID DENC
    mocked_text = """
    CLIENTE: Renting de Antioquia EICE
    DOCUMENTO: 900285704
    TELEFONO: 3016904190
    DIRECCION: Cra 43A NO.19127, Medellín / Colombia

    2025-10-15 / 19:26:06
    2025-10-15 / 19:26:33
    Vencimiento:
    2025-10-15

    Método de pago: Contado
    Medio de pago: Trns.Debito

    Código  Descripción  Unidad  Cant  Unitario  Total
    1       Placa: GEM560::Peaje: AMAGA::Categoria: 1::Fecha del paso: 16/10/2025 00:19:00  94  1,00  16.100,00  16.100,00

    DENC-00012345
    """.strip()
    monkeypatch.setattr(ZipInvoiceExtractor, "pdf_to_text", lambda *a, **k: mocked_text)

    # Ejecuta workflow end-to-end
    result = run_pipeline(
        extract_dir=extract,
        json_dir=out_json,
        text_dir=out_text,
        index_csv=index_csv,
        lang="spa",
        dpi=150,
    )
    ex = result["extract"]

    # Verificaciones básicas
    assert ex.zips == 2
    assert ex.docs_json == 2
    assert ex.errors == 0

    # Verifica JSONs y series DEFL/DENC
    json_files = list(out_json.rglob("*.json"))
    assert len(json_files) == 2

    ids = []
    for jf in json_files:
        data = json.loads(jf.read_text(encoding="utf-8"))
        doc = data.get("document", {})
        assert doc.get("document_id")
        ids.append(doc["document_id"])

    assert any(i.startswith("DEFL-") for i in ids)
    assert any(i.startswith("DENC-") for i in ids)
