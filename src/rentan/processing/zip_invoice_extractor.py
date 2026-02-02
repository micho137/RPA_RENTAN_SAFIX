from __future__ import annotations

import csv
import logging
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, List, Dict

# Fallbacks opcionales
try:
    from pdfminer.high_level import extract_text as pdf_extract_text
except Exception:
    pdf_extract_text = None

try:
    from pdf2image import convert_from_path
except Exception:
    convert_from_path = None

try:
    import pytesseract
except Exception:
    pytesseract = None

from src.rentan.processing.doc_id import normalize_id, parse_any_id, split_parts


NS = {
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "inv": "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2",
}


def _clean(s: Optional[str]) -> str:
    return (s or "").strip()


def _to_money_int(num: Optional[str]) -> int:
    """
    Convierte valores monetarios a entero (pesos), SIN agregar ceros.

    Soporta:
      - "19.100"       -> 19100      (miles con punto)
      - "19.100,00"    -> 19100      (miles '.' decimal ',')
      - "19,100.00"    -> 19100      (miles ',' decimal '.')
      - "19100.00"     -> 19100      (decimal con punto)
      - "19100,00"     -> 19100      (decimal con coma)
      - "1.910.000"    -> 1910000
      - "1,910,000.00" -> 1910000
    """
    if not num:
        return 0

    s = str(num).strip().replace(" ", "").replace("\u00A0", "")
    if not s:
        return 0

    # Mantener solo dígitos, separadores y signo
    s = re.sub(r"[^\d,.\-+]", "", s)

    # Signo (por si aparece)
    sign = -1 if s.startswith("-") else 1
    if s[:1] in "+-":
        s = s[1:]

    # Caso A: tiene '.' y ',' -> el último separador es el decimal
    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            # ES/CO: miles '.', decimal ','
            s = s.replace(".", "")
            s = s.split(",", 1)[0]  # cortar decimales
        else:
            # US: miles ',', decimal '.'
            s = s.replace(",", "")
            s = s.split(".", 1)[0]  # cortar decimales

    # Caso B: solo coma
    elif "," in s:
        parts = s.split(",")
        # si es NNNN,dd (2 dígitos) => decimal, cortar
        if len(parts) == 2 and len(parts[1]) == 2:
            s = parts[0]
        else:
            # si no, asumir miles con coma
            s = s.replace(",", "")

    # Caso C: solo punto
    elif "." in s:
        parts = s.split(".")
        # si es NNNN.dd (2 dígitos) => decimal, cortar (EVITA 19100.00 -> 1910000)
        if len(parts) == 2 and len(parts[1]) == 2:
            s = parts[0]
        else:
            # si no, asumir miles con punto (o varios puntos)
            s = s.replace(".", "")

    # Solo dígitos
    s = re.sub(r"[^\d]", "", s)
    if not s:
        return 0

    return sign * int(s)


def _to_qty_float(num: Optional[str]) -> float:
    """
    Convierte cantidades (pueden tener decimales) a float de forma robusta.
    - "1.234,5" -> 1234.5
    - "1,234.5" -> 1234.5
    - "0,5"     -> 0.5
    """
    if not num:
        return 0.0

    s = str(num).strip().replace(" ", "").replace("\u00A0", "")
    if not s:
        return 0.0

    s = re.sub(r"[^\d,.\-+]", "", s)

    try:
        if "." in s and "," in s:
            # último separador es decimal
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
        return float(s)
    except ValueError:
        return 0.0


def _remove_hyphen_in_id(doc_id: Optional[str]) -> Optional[str]:
    """
    Convierte DEFL-123 -> DEFL123 y DENC-123 -> DENC123
    Si ya está sin guion, lo deja igual.
    """
    if not doc_id:
        return doc_id
    return doc_id.replace("-", "").strip()


@dataclass
class ExtractResult:
    zips: int = 0
    docs_json: int = 0
    errors: int = 0


class ZipInvoiceExtractor:
    """
    Recorre .zip en download_dir, extrae a extract_root/<zipname>/,
    y para cada documento intenta:
      1) UBL XML (directo o invoice embebido en AttachedDocument) -> JSON
      2) Si NO hay XML, PDF->texto; si no hay texto, OCR -> JSON

    Guarda:
      - JSON por documento en out_json_root
        * Si existe document_id => <out_json_root>/<document_id>.json (evita duplicados)
        * Si no existe => replica jerarquía (fallback)
      - TXT (si vino de PDF/OCR) en out_text_root (opcional)
      - índice CSV (opcional)
    """

    def __init__(
        self,
        download_dir: Path,
        extract_root: Path,
        out_json_root: Path,
        out_text_root: Optional[Path] = None,
        lang: str = "spa",
        dpi: int = 300,
        logger: Optional[logging.Logger] = None,
        overwrite_json: bool = True,
    ):
        self.download_dir = Path(download_dir)
        self.extract_root = Path(extract_root)
        self.out_json_root = Path(out_json_root)
        self.out_text_root = Path(out_text_root) if out_text_root else None
        self.lang = lang
        self.dpi = dpi
        self.overwrite_json = overwrite_json
        self.log = logger or logging.getLogger("zip_invoice_extractor")

        self.extract_root.mkdir(parents=True, exist_ok=True)
        self.out_json_root.mkdir(parents=True, exist_ok=True)
        if self.out_text_root:
            self.out_text_root.mkdir(parents=True, exist_ok=True)

    # ------------------ ZIP ------------------
    def iter_zip_files(self) -> Iterable[Path]:
        for p in sorted(self.download_dir.rglob("*.zip")):
            if p.is_file():
                yield p

    def extract_zip(self, zip_path: Path) -> Path:
        target_dir = self.extract_root / zip_path.stem
        target_dir.mkdir(parents=True, exist_ok=True)
        if any(target_dir.iterdir()):
            self.log.info(f"Skip extract (exists): {target_dir}")
            return target_dir
        self.log.info(f"Extract: {zip_path} -> {target_dir}")
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(target_dir)
        return target_dir

    # ------------------ XML (UBL primero) ------------------
    def parse_invoice_from_ubl(self, xml_path: Path) -> Optional[Dict]:
        """
        Soporta <Invoice> directo o AttachedDocument con <Invoice> embebido en Description.
        """
        try:
            root = ET.parse(xml_path).getroot()
        except Exception as e:
            self.log.warning(f"Invalid XML {xml_path.name}: {e}")
            return None

        # a) ¿Invoice directo?
        if root.tag.endswith("Invoice"):
            invoice = root
        else:
            # b) AttachedDocument con <Invoice> en Description (CDATA)
            desc = root.find(".//{urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2}Description")
            if desc is None or not desc.text or "<Invoice" not in desc.text:
                return None
            try:
                invoice = ET.fromstring(desc.text)
            except Exception as e:
                self.log.warning(f"Embedded Invoice parse failed in {xml_path.name}: {e}")
                return None

        # Emisor / Cliente
        emisor_nombre = _clean(invoice.findtext(".//cac:AccountingSupplierParty/cac:Party/cac:PartyName/cbc:Name", namespaces=NS))
        emisor_nit = _clean(invoice.findtext(".//cac:AccountingSupplierParty/cac:PartyTaxScheme/cbc:CompanyID", namespaces=NS))

        cliente_nombre = _clean(invoice.findtext(".//cac:AccountingCustomerParty/cac:Party/cac:PartyName/cbc:Name", namespaces=NS))
        cliente_nit = _clean(invoice.findtext(".//cac:AccountingCustomerParty/cac:PartyTaxScheme/cbc:CompanyID", namespaces=NS))
        cliente_tel = _clean(invoice.findtext(".//cac:AccountingCustomerParty/cac:Party/cac:Contact/cbc:Telephone", namespaces=NS))

        linea_dir = _clean(invoice.findtext(".//cac:AccountingCustomerParty/cac:Party/cac:PhysicalLocation/cac:Address/cac:AddressLine/cbc:Line", namespaces=NS))
        ciudad = _clean(invoice.findtext(".//cac:AccountingCustomerParty/cac:Party/cac:PhysicalLocation/cac:Address/cbc:CityName", namespaces=NS))
        pais = _clean(invoice.findtext(".//cac:AccountingCustomerParty/cac:Party/cac:PhysicalLocation/cac:Address/cac:Country/cbc:Name", namespaces=NS))
        direccion = ", ".join([p for p in [linea_dir, f"{ciudad} / {pais}" if ciudad and pais else ciudad or pais] if p])

        # IDs / fechas
        doc_id_raw = _clean(invoice.findtext("cbc:ID", namespaces=NS))
        document_id = normalize_id(doc_id_raw) or doc_id_raw or None
        document_id = _remove_hyphen_in_id(document_id)

        serie, numero = split_parts(document_id or "")
        cufe = _clean(invoice.findtext("cbc:UUID", namespaces=NS))
        issue_date = _clean(invoice.findtext("cbc:IssueDate", namespaces=NS))
        issue_time = _clean(invoice.findtext("cbc:IssueTime", namespaces=NS))
        due_date = _clean(invoice.findtext("cac:PaymentMeans/cbc:PaymentDueDate", namespaces=NS))
        moneda = _clean(invoice.findtext("cbc:DocumentCurrencyCode", namespaces=NS)) or "COP"

        # Totales (DINERO -> int)
        subtotal = _to_money_int(_clean(invoice.findtext("cac:LegalMonetaryTotal/cbc:LineExtensionAmount", namespaces=NS)))
        total = _to_money_int(_clean(invoice.findtext("cac:LegalMonetaryTotal/cbc:PayableAmount", namespaces=NS)))

        # Ítems (mínimo 1)
        line = invoice.find(".//cac:InvoiceLine", namespaces=NS)
        codigo = _clean(line.findtext("cbc:ID", namespaces=NS)) if line is not None else ""
        desc = _clean(line.findtext("cac:Item/cbc:Description", namespaces=NS)) if line is not None else ""
        unidad = ""
        cantidad = 0.0  # cantidad puede ser decimal
        unitario = 0    # dinero int
        total_linea = 0 # dinero int

        if line is not None:
            qty = line.find("cbc:InvoicedQuantity", namespaces=NS)
            if qty is not None:
                unidad = qty.attrib.get("unitCode", "")
                cantidad = _to_qty_float(qty.text or "0")

            unitario = _to_money_int(_clean(line.findtext("cac:Price/cbc:PriceAmount", namespaces=NS)))
            total_linea = _to_money_int(_clean(line.findtext("cbc:LineExtensionAmount", namespaces=NS)))

        data = {
            "document": {
                "serie": serie or None,
                "numero": numero or None,
                "document_id": document_id,
            },
            "cufe": cufe or None,
            "cliente": {
                "nombre": cliente_nombre or None,
                "nit": cliente_nit or None,
                "telefono": cliente_tel or None,
                "direccion": direccion or None,
            },
            "emisor": {"nombre": emisor_nombre or None, "nit": emisor_nit or None},
            "fechas": {
                "venta": f"{issue_date}T{issue_time}" if issue_date and issue_time else None,
                "expedicion": None,
                "vencimiento": due_date or None,
            },
            "pago": {"metodo": None, "medio": None},
            "moneda": moneda,
            "totales": {"subtotal": subtotal, "total": total},
            "detalle": [
                {
                    "codigo": codigo or None,
                    "descripcion": desc or None,
                    "unidad": unidad or None,
                    "cantidad": cantidad,
                    "unitario": unitario,
                    "total": total_linea,
                }
            ],
        }
        return data

    # ------------------ PDF→texto/OCR ------------------
    def pdf_to_text(self, pdf_path: Path) -> str:
        # Texto embebido
        if pdf_extract_text is not None:
            try:
                txt = pdf_extract_text(str(pdf_path)) or ""
                if txt and len(txt) > 30:
                    return txt
            except Exception as e:
                self.log.warning(f"pdfminer failed {pdf_path.name}: {e}")

        # Fallback OCR
        if convert_from_path is None or pytesseract is None:
            raise RuntimeError("For PDF OCR install: pdf2image, pytesseract and system deps (poppler, tesseract).")

        pages = convert_from_path(str(pdf_path), dpi=self.dpi)
        chunks: List[str] = []
        for page in pages:
            if page.mode not in ("RGB", "L"):
                page = page.convert("RGB")
            chunks.append(pytesseract.image_to_string(page, lang=self.lang))
        return "\n\n".join(chunks)

    # ------------------ Regex desde texto (PDF/OCR) ------------------
    def parse_from_text(self, txt: str) -> Dict:
        base = {
            "document": {"serie": None, "numero": None, "document_id": None},
            "cufe": None,
            "cliente": {"nombre": None, "nit": None, "telefono": None, "direccion": None},
            "emisor": {"nombre": None, "nit": None},
            "fechas": {"venta": None, "expedicion": None, "vencimiento": None},
            "pago": {"metodo": None, "medio": None},
            "moneda": "COP",
            "totales": {"subtotal": 0, "total": 0},
            "detalle": [],
        }

        # Doc ID (DEFL/DENC con o sin guion)
        found = parse_any_id(txt)
        if found:
            serie, numero, norm = found
            norm = _remove_hyphen_in_id(norm)
            base["document"]["serie"] = serie or None
            base["document"]["numero"] = numero or None
            base["document"]["document_id"] = norm

        # Cliente / Doc / Tel / Dirección (ajusta etiquetas si cambian)
        m = re.search(r"CLIENTE:\s*(.+)", txt, re.IGNORECASE)
        if m:
            base["cliente"]["nombre"] = _clean(m.group(1))

        m = re.search(r"DOCUMENTO:\s*([\d.]+)", txt, re.IGNORECASE)
        if m:
            base["cliente"]["nit"] = _clean(m.group(1)).replace(".", "")

        m = re.search(r"TELEFONO:\s*([\d\s]+)", txt, re.IGNORECASE)
        if m:
            base["cliente"]["telefono"] = _clean(m.group(1)).replace(" ", "")

        m = re.search(r"DIRECCION:\s*(.+)", txt, re.IGNORECASE)
        if m:
            base["cliente"]["direccion"] = _clean(m.group(1))

        # Fechas (dos timestamps: venta y expedición)
        stamps = re.findall(r"(\d{4}-\d{2}-\d{2})\s*/\s*(\d{2}:\d{2}:\d{2})", txt)
        if len(stamps) >= 1:
            base["fechas"]["venta"] = f"{stamps[0][0]}T{stamps[0][1]}"
        if len(stamps) >= 2:
            base["fechas"]["expedicion"] = f"{stamps[1][0]}T{stamps[1][1]}"

        m = re.search(r"Vencimiento:\s*\n?\s*(\d{4}-\d{2}-\d{2})", txt, re.IGNORECASE)
        if m:
            base["fechas"]["vencimiento"] = _clean(m.group(1))

        # Pago
        m = re.search(r"Método de pago:\s*(.+)", txt, re.IGNORECASE)
        if m:
            base["pago"]["metodo"] = _clean(m.group(1))

        m = re.search(r"Medio de pago:\s*(.+)", txt, re.IGNORECASE)
        if m:
            base["pago"]["medio"] = _clean(m.group(1))

        # Tabla del ítem (una línea)
        it = re.search(
            r"Código\s+Descripción\s+Unidad\s+Cant\s+Unitario\s+Total\s+(\d+)\s+(.+?)\s+(\S+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)",
            txt,
            re.IGNORECASE | re.DOTALL,
        )
        if it:
            cantidad = _to_qty_float(it.group(4))
            unitario = _to_money_int(it.group(5))
            total_linea = _to_money_int(it.group(6))

            base["detalle"] = [{
                "codigo": _clean(it.group(1)),
                "descripcion": _clean(it.group(2)),
                "unidad": _clean(it.group(3)),
                "cantidad": cantidad,
                "unitario": unitario,
                "total": total_linea,
            }]
            base["totales"]["subtotal"] = total_linea
            base["totales"]["total"] = total_linea

        return base

    # ------------------ Guardado ------------------
    def save_json_for(self, source_path: Path, data: Dict) -> Path:
        """
        Evita duplicación:
          - Si hay document_id => <out_json_root>/<document_id>.json
          - Si no hay => replica jerarquía del extract_root
        """
        import json

        doc_id = ((data.get("document") or {}).get("document_id")) or ""
        doc_id = str(doc_id).strip()

        if doc_id:
            out_json = self.out_json_root / f"{doc_id}.json"
        else:
            rel = source_path.relative_to(self.extract_root)
            out_json = (self.out_json_root / rel).with_suffix(".json")

        out_json.parent.mkdir(parents=True, exist_ok=True)

        if out_json.exists() and not self.overwrite_json:
            self.log.info(f"Skip JSON (exists): {out_json}")
            return out_json

        out_json.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return out_json

    def save_txt_for(self, source_path: Path, text: str) -> Optional[Path]:
        if not self.out_text_root:
            return None
        rel = source_path.relative_to(self.extract_root)
        out_txt = (self.out_text_root / rel).with_suffix(".txt")
        out_txt.parent.mkdir(parents=True, exist_ok=True)
        out_txt.write_text(text, encoding="utf-8", errors="ignore")
        return out_txt

    # ------------------ Proceso principal ------------------
    def process_all(self, index_csv: Optional[Path] = None) -> ExtractResult:
        res = ExtractResult()
        index_rows: List[Dict] = []

        for zip_path in self.iter_zip_files():
            res.zips += 1
            try:
                extracted_dir = self.extract_zip(zip_path)
            except Exception as e:
                self.log.error(f"Extract error {zip_path.name}: {e}")
                res.errors += 1
                continue

            # 1) Intentar XMLs primero
            xmls = [p for p in extracted_dir.rglob("*.xml") if p.is_file()]
            any_xml_ok = False
            for xml in xmls:
                data = self.parse_invoice_from_ubl(xml)
                if not data:
                    continue
                any_xml_ok = True
                out_json = self.save_json_for(xml, data)
                res.docs_json += 1
                index_rows.append({
                    "zip": zip_path.name,
                    "source": str(xml),
                    "json": str(out_json),
                    "mode": "XML",
                    "document_id": ((data.get("document") or {}).get("document_id")) or "",
                    "serie": ((data.get("document") or {}).get("serie")) or "",
                })

            if any_xml_ok:
                continue

            # 2) Si no hay XML válido, intentar PDFs
            pdfs = [p for p in extracted_dir.rglob("*.pdf") if p.is_file()]
            for pdf in pdfs:
                try:
                    text = self.pdf_to_text(pdf)
                    self.save_txt_for(pdf, text)
                    data = self.parse_from_text(text)
                    out_json = self.save_json_for(pdf, data)
                    res.docs_json += 1
                    index_rows.append({
                        "zip": zip_path.name,
                        "source": str(pdf),
                        "json": str(out_json),
                        "mode": "PDF",
                        "document_id": ((data.get("document") or {}).get("document_id")) or "",
                        "serie": ((data.get("document") or {}).get("serie")) or "",
                    })
                except Exception as e:
                    self.log.error(f"PDF/OCR fail {pdf.name}: {e}")
                    res.errors += 1
                    continue

        if index_csv:
            try:
                index_csv.parent.mkdir(parents=True, exist_ok=True)
                with open(index_csv, "w", newline="", encoding="utf-8") as f:
                    writer = csv.DictWriter(
                        f,
                        fieldnames=["zip", "source", "json", "mode", "document_id", "serie"],
                    )
                    writer.writeheader()
                    writer.writerows(index_rows)
            except Exception as e:
                self.log.warning(f"Index write failed: {e}")

        return res
