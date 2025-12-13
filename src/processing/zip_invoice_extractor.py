from __future__ import annotations
import csv
import logging
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Tuple, List, Dict
import re
import xml.etree.ElementTree as ET

from PIL import Image

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

from src.processing.doc_id import normalize_id, parse_any_id, split_parts

# Namespaces UBL base
NS = {
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "inv": "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2",
}

# NEW: namespace estructuras DIAN dentro de UBL
NS_STS = {
    "sts": "dian:gov:co:facturaelectronica:Structures-2-1",
}

def _clean(s: Optional[str]) -> str:
    return (s or "").strip()

def _to_float(num: str) -> float:
    """
    Convierte cadenas con formatos:
      - '16100.00'          (XML estándar, punto decimal)
      - '16.100,00'         (PDF español, punto miles + coma decimal)
      - '16100'             (entero)
      - '16,100.00'         (menos frecuente: coma miles + punto decimal)
    """
    s = (num or "").strip()
    if not s:
        return 0.0

    # Caso con punto y coma: decidir cuál es decimal
    if "." in s and "," in s:
        # Si la coma está DESPUÉS del último punto → formato '16.100,00'
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")  # 16.100,00 -> 16100.00
        else:
            # Formato tipo '16,100.00' → coma de miles, punto decimal
            s = s.replace(",", "")
    else:
        # Sólo coma: la tratamos como decimal (estilo español)
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        # Sólo punto: lo dejamos tal cual (estilo anglosajón / XML)
        # Ningún separador: entero plano

    try:
        return float(s)
    except Exception:
        return 0.0


# NEW: helpers para fusionar XML+PDF
def _merge_scalar(primary, secondary):
    """
    Si primary tiene valor (no None/""), lo conserva;
    si está vacío, usa secondary.
    """
    if primary not in (None, ""):
        return primary
    return secondary

def _merge_dict(primary: Optional[Dict], secondary: Optional[Dict]) -> Dict:
    """
    Combina dos dicts, conservando los valores de primary cuando existen
    y usando secondary solo para rellenar faltantes.
    """
    out: Dict = dict(primary or {})
    for k, v in (secondary or {}).items():
        if k not in out or out[k] in (None, "", []):
            out[k] = v
    return out

def _format_datetime(date_str: str, time_str: str) -> Optional[str]:
    """
    Une fecha + hora en formato 'YYYY-MM-DD HH:MM:SS', eliminando el offset
    de zona horaria si viene (ej. '05:40:31-05:00' -> '05:40:31').
    """
    d = _clean(date_str)
    t = _clean(time_str)
    if not d or not t:
        return None
    # Quitar parte de zona horaria (+/-HH:MM)
    t_clean = re.split(r"[+-]", t)[0]  # '05:40:31-05:00' -> '05:40:31'
    return f"{d} {t_clean}"

def _split_descripcion(desc: str) -> Dict[str, str]:
    """
    A partir de 'Placa: HYT426::Peaje: AMAGA::Categoria: 1::Fecha del paso: 29/10/2025 10:28:11'
    devuelve:
      {
        "placa": "HYT426",
        "peaje": "AMAGA",
        "categoria": "1",
        "fecha_del_paso": "29/10/2025 10:28:11"
      }
    """
    result: Dict[str, str] = {}
    if not desc:
        return result

    parts = [p.strip() for p in desc.split("::") if p.strip()]
    for part in parts:
        if ":" not in part:
            continue
        key, value = part.split(":", 1)
        k = _clean(key).lower().replace(" ", "_")
        v = _clean(value)
        if k:
            result[k] = v
    return result


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
      - .json por documento (en out_json_root, replicando jerarquía)
      - .txt (si vino de PDF/OCR) en out_text_root (opcional)
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
    ):
        self.download_dir = Path(download_dir)
        self.extract_root = Path(extract_root)
        self.out_json_root = Path(out_json_root)
        self.out_text_root = Path(out_text_root) if out_text_root else None
        self.lang = lang
        self.dpi = dpi
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
        Enriquecido con:
          - datos DIAN (resolución, rango, vigencia, QR, etc.)
          - mapeo básico de método de pago desde PaymentMeansCode
          - bloque 'tributario' con actividades económicas.
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
            desc = root.find(
                ".//{urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2}Description"
            )
            if desc is None or not desc.text or "<Invoice" not in desc.text:
                return None
            try:
                invoice = ET.fromstring(desc.text)
            except Exception as e:
                self.log.warning(f"Embedded Invoice parse failed in {xml_path.name}: {e}")
                return None

        # ---------- Emisor / Cliente ----------
        emisor_nombre = _clean(
            invoice.findtext(
                ".//cac:AccountingSupplierParty/cac:Party/cac:PartyName/cbc:Name",
                namespaces=NS,
            )
        )
        emisor_nit = _clean(
            invoice.findtext(
                ".//cac:AccountingSupplierParty/cac:Party/cac:PartyTaxScheme/cbc:CompanyID",
                namespaces=NS,
            )
        )

        cliente_nombre = _clean(
            invoice.findtext(
                ".//cac:AccountingCustomerParty/cac:Party/cac:PartyName/cbc:Name",
                namespaces=NS,
            )
        )
        cliente_nit = _clean(
            invoice.findtext(
                ".//cac:AccountingCustomerParty/cac:Party/cac:PartyTaxScheme/cbc:CompanyID",
                namespaces=NS,
            )
        )
        cliente_tel = _clean(
            invoice.findtext(
                ".//cac:AccountingCustomerParty/cac:Party/cac:Contact/cbc:Telephone",
                namespaces=NS,
            )
        )

        linea_dir = _clean(
            invoice.findtext(
                ".//cac:AccountingCustomerParty/cac:Party/cac:PhysicalLocation/"
                "cac:Address/cac:AddressLine/cbc:Line",
                namespaces=NS,
            )
        )
        ciudad = _clean(
            invoice.findtext(
                ".//cac:AccountingCustomerParty/cac:Party/cac:PhysicalLocation/"
                "cac:Address/cbc:CityName",
                namespaces=NS,
            )
        )
        pais = _clean(
            invoice.findtext(
                ".//cac:AccountingCustomerParty/cac:Party/cac:PhysicalLocation/"
                "cac:Address/cac:Country/cbc:Name",
                namespaces=NS,
            )
        )
        direccion = ", ".join(
            [
                p
                for p in [
                    linea_dir,
                    f"{ciudad} / {pais}" if ciudad and pais else ciudad or pais,
                ]
                if p
            ]
        )

        # ---------- IDs / fechas ----------
        doc_id_raw = _clean(invoice.findtext("cbc:ID", namespaces=NS))
        document_id = normalize_id(doc_id_raw) or doc_id_raw or None
        serie, numero = split_parts(document_id or "")
        cufe = _clean(invoice.findtext("cbc:UUID", namespaces=NS))
        issue_date = _clean(invoice.findtext("cbc:IssueDate", namespaces=NS))
        issue_time = _clean(invoice.findtext("cbc:IssueTime", namespaces=NS))
        due_date = _clean(
            invoice.findtext("cac:PaymentMeans/cbc:PaymentDueDate", namespaces=NS)
        )
        moneda = _clean(
            invoice.findtext("cbc:DocumentCurrencyCode", namespaces=NS)
        ) or "COP"

        # ---------- Totales ----------
        subtotal = _to_float(
            _clean(
                invoice.findtext(
                    "cac:LegalMonetaryTotal/cbc:LineExtensionAmount", namespaces=NS
                )
            )
        )
        total = _to_float(
            _clean(
                invoice.findtext(
                    "cac:LegalMonetaryTotal/cbc:PayableAmount", namespaces=NS
                )
            )
        )

        # ---------- Ítems (mín. uno) ----------
        line = invoice.find(".//cac:InvoiceLine", namespaces=NS)
        codigo = _clean(line.findtext("cbc:ID", namespaces=NS)) if line is not None else ""
        desc_raw = _clean(
            line.findtext("cac:Item/cbc:Description", namespaces=NS)
        ) if line is not None else ""
        desc_parsed = _split_descripcion(desc_raw)
        unidad = ""
        cantidad = 0.0
        unitario = 0.0
        total_linea = 0.0
        if line is not None:
            qty = line.find("cbc:InvoicedQuantity", namespaces=NS)
            if qty is not None:
                unidad = qty.attrib.get("unitCode", "")
                cantidad = _to_float(qty.text or "0")
            unitario = _to_float(
                _clean(line.findtext("cac:Price/cbc:PriceAmount", namespaces=NS))
            )
            total_linea = _to_float(
                _clean(line.findtext("cbc:LineExtensionAmount", namespaces=NS))
            )

        # ---------- Método de pago desde PaymentMeansCode ----------
        payment_means_code = _clean(
            invoice.findtext("cac:PaymentMeans/cbc:PaymentMeansCode", namespaces=NS)
        )
        payment_id = _clean(
            invoice.findtext("cac:PaymentMeans/cbc:ID", namespaces=NS)
        )

        metodo_pago = None
        # Mapeo muy básico; ajústalo a tu tabla DIAN real
        if payment_means_code == "31":
            metodo_pago = "Contado"
        elif payment_means_code in {"2", "4"}:
            metodo_pago = "Crédito"

        # ---------- Datos DIAN desde sts:DianExtensions ----------
        dian_resolucion = None
        dian_rango_desde = None
        dian_rango_hasta = None
        dian_fec_ini = None
        dian_fec_fin = None
        dian_qr = None
        dian_sw_id = None
        dian_sw_nit = None
        dian_auth_provider = None

        dian_ext = invoice.find(".//sts:DianExtensions", namespaces=NS_STS)
        if dian_ext is not None:
            inv_control = dian_ext.find("sts:InvoiceControl", namespaces=NS_STS)
            if inv_control is not None:
                dian_resolucion = _clean(
                    inv_control.findtext("sts:InvoiceAuthorization", namespaces=NS_STS)
                )
                dian_rango_desde = _clean(
                    inv_control.findtext(
                        "sts:AuthorizedInvoices/sts:From", namespaces=NS_STS
                    )
                )
                dian_rango_hasta = _clean(
                    inv_control.findtext(
                        "sts:AuthorizedInvoices/sts:To", namespaces=NS_STS
                    )
                )
                dian_fec_ini = _clean(
                    inv_control.findtext(
                        "sts:AuthorizationPeriod/cbc:StartDate",
                        namespaces={**NS, **NS_STS},
                    )
                )
                dian_fec_fin = _clean(
                    inv_control.findtext(
                        "sts:AuthorizationPeriod/cbc:EndDate",
                        namespaces={**NS, **NS_STS},
                    )
                )

            sw = dian_ext.find("sts:SoftwareProvider", namespaces=NS_STS)
            if sw is not None:
                dian_sw_nit = _clean(
                    sw.findtext("sts:ProviderID", namespaces=NS_STS)
                )
                dian_sw_id = _clean(sw.findtext("sts:SoftwareID", namespaces=NS_STS))

            authp = dian_ext.find("sts:AuthorizationProvider", namespaces=NS_STS)
            if authp is not None:
                dian_auth_provider = _clean(
                    authp.findtext("sts:AuthorizationProviderID", namespaces=NS_STS)
                )

            dian_qr = _clean(dian_ext.findtext("sts:QRCode", namespaces=NS_STS))

        # ---------- Tributario (actividades) ----------
        tributario = {
            "responsable_iva": None,
            "autoretenedor_renta": None,
            "autoretenedor_ica": None,
            "gran_contribuyente": None,
            "actividades": [],
        }
        act_str = _clean(
            invoice.findtext(
                ".//cac:AccountingSupplierParty/cac:Party/cbc:IndustryClassificationCode",
                namespaces=NS,
            )
        )
        if act_str:
            # En el ejemplo viene "6201;7112;6202"
            tributario["actividades"] = [
                a for a in re.split(r"[;\s]+", act_str) if a
            ]

        # ---------- Bloque DIAN ----------
        dian = {
            "resolucion": dian_resolucion,
            "rango_desde": dian_rango_desde,
            "rango_hasta": dian_rango_hasta,
            "fec_ini_autoriz": dian_fec_ini,
            "fec_fin_autoriz": dian_fec_fin,
            "qr": dian_qr,
            "software_nit": dian_sw_nit,
            "software_id": dian_sw_id,
            "authorization_provider": dian_auth_provider,
            # Estos campos pueden completarse luego desde ApplicationResponse/PDF
            "estado_validacion": None,
            "mensaje_validacion": None,
        }

        data = {
            "document": {
                "serie": serie,
                "numero": numero,
                "document_id": document_id,  # DEFL-######## o DENC-########
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
                "venta": _format_datetime(issue_date, issue_time),
                "expedicion": None,  # se completará desde PDF
                "vencimiento": due_date or None,
            },
            "pago": {
                "metodo": metodo_pago,
                "medio": None,        # complementable desde PDF
                "payment_id": payment_id or None,
            },
            "moneda": moneda,
            "totales": {"subtotal": subtotal, "total": total},
            "detalle": [
                {
                    "codigo": codigo or None,
                    "descripcion_raw": desc_raw or None,
                    "descripcion": desc_parsed or None,
                    "unidad": unidad or None,
                    "cantidad": cantidad,
                    "unitario": unitario,
                    "total": total_linea,
                }
            ],
            "tributario": tributario,
            "dian": dian,
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
        for i, page in enumerate(pages, start=1):
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
            "totales": {"subtotal": 0.0, "total": 0.0},
            "detalle": [],
            # NEW
            "tributario": {
                "responsable_iva": None,
                "autoretenedor_renta": None,
                "autoretenedor_ica": None,
                "gran_contribuyente": None,
                "actividades": [],
            },
            "dian": {
                "resolucion": None,
                "rango_desde": None,
                "rango_hasta": None,
                "fec_ini_autoriz": None,
                "fec_fin_autoriz": None,
                "qr": None,
                "software_nit": None,
                "software_id": None,
                "authorization_provider": None,
                "estado_validacion": None,
                "mensaje_validacion": None,
            },
        }

        # Doc ID (DEFL/DENC con o sin guion)
        found = parse_any_id(txt)
        if found:
            serie, numero, norm = found
            base["document"]["serie"] = serie
            base["document"]["numero"] = numero
            base["document"]["document_id"] = norm

        # Cliente / Doc / Tel / Dirección
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
            base["fechas"]["venta"] = f"{stamps[0][0]} {stamps[0][1]}"
        if len(stamps) >= 2:
            base["fechas"]["expedicion"] = f"{stamps[1][0]} {stamps[1][1]}"
        m = re.search(r"Vencimiento:\s*\n?\s*(\d{4}-\d{2}-\d{2})", txt, re.IGNORECASE)
        if m:
            base["fechas"]["vencimiento"] = _clean(m.group(1))

        # Pago: intentar capturar bloque método + medio en una sola pasada
        mp = re.search(
            r"M[eé]todo de pago:\s*([^\n\r]+).*?Medio de pago:\s*([^\n\r]+)",
            txt,
            re.IGNORECASE | re.DOTALL,
        )
        if mp:
            base["pago"]["metodo"] = _clean(mp.group(1))
            base["pago"]["medio"] = _clean(mp.group(2))
        else:
            # Fallback simple por si el layout cambia
            m = re.search(r"M[eé]todo de pago:\s*([^\n\r]+)", txt, re.IGNORECASE)
            if m:
                base["pago"]["metodo"] = _clean(m.group(1))
            m = re.search(r"Medio de pago:\s*([^\n\r]+)", txt, re.IGNORECASE)
            if m:
                base["pago"]["medio"] = _clean(m.group(1))

        # CUFE/CUDE
        m = re.search(r"CUDE:\s*([0-9a-fA-F]+)", txt, re.IGNORECASE)
        if m:
            base["cufe"] = _clean(m.group(1))

        # DIAN: numeración, resolución, vigencia (si aparecen en el PDF)
        m = re.search(
            r"NUMERACION AUTORIZADA\s+DEL\s+([A-Z]+)-?(\d+)\s+AL\s+([A-Z]+)-?(\d+)",
            txt,
            re.IGNORECASE,
        )
        if m:
            base["dian"]["rango_desde"] = _clean(m.group(2))
            base["dian"]["rango_hasta"] = _clean(m.group(4))
        m = re.search(r"RESOLUCION:\s*([0-9]+)", txt, re.IGNORECASE)
        if m:
            base["dian"]["resolucion"] = _clean(m.group(1))
        m = re.search(
            r"DESDE:\s*(\d{4}-\d{2}-\d{2})\s+HASTA:\s*(\d{4}-\d{2}-\d{2})",
            txt,
            re.IGNORECASE,
        )
        if m:
            base["dian"]["fec_ini_autoriz"] = _clean(m.group(1))
            base["dian"]["fec_fin_autoriz"] = _clean(m.group(2))

        # Tributario flags
        if re.search(r"IVA\s*-\s*Responsables", txt, re.IGNORECASE):
            base["tributario"]["responsable_iva"] = "Responsables"
        elif re.search(r"IVA\s*-\s*No\s+responsables", txt, re.IGNORECASE):
            base["tributario"]["responsable_iva"] = "No responsables"

        if re.search(r"No\s+somos\s+autoretenedores\b", txt, re.IGNORECASE):
            base["tributario"]["autoretenedor_renta"] = False
        elif re.search(r"Somos\s+autoretenedores\b(?!\s*ICA)", txt, re.IGNORECASE):
            base["tributario"]["autoretenedor_renta"] = True

        if re.search(r"Somos\s+autoretenedores\s+ICA", txt, re.IGNORECASE):
            base["tributario"]["autoretenedor_ica"] = True
        elif re.search(r"No\s+somos\s+autoretenedores\s+ICA", txt, re.IGNORECASE):
            base["tributario"]["autoretenedor_ica"] = False

        if re.search(r"No\s+somos\s+grandes\s+contribuyentes", txt, re.IGNORECASE):
            base["tributario"]["gran_contribuyente"] = False
        elif re.search(r"Somos\s+grandes\s+contribuyentes", txt, re.IGNORECASE):
            base["tributario"]["gran_contribuyente"] = True

        m = re.search(r"Actividades\s+([\d\s]+)", txt, re.IGNORECASE)
        if m:
            acts = [a for a in re.split(r"\s+", _clean(m.group(1))) if a]
            base["tributario"]["actividades"] = acts

        # Tabla del ítem (una línea)
        it = re.search(
            r"C[oó]digo\s+Descripci[oó]n\s+Unidad\s+Cant\s+Unitario\s+Total\s+"
            r"(\d+)\s+(.+?)\s+(\S+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)",
            txt,
            re.IGNORECASE | re.DOTALL,
        )
        if it:
            desc_raw = _clean(it.group(2))
            desc_parsed = _split_descripcion(desc_raw)
            base["detalle"] = [
                {
                    "codigo": _clean(it.group(1)),
                    "descripcion_raw": desc_raw or None,
                    "descripcion": desc_parsed or None,
                    "unidad": _clean(it.group(3)),
                    "cantidad": _to_float(it.group(4)),
                    "unitario": _to_float(it.group(5)),
                    "total": _to_float(it.group(6)),
                }
            ]
            base["totales"]["subtotal"] = base["detalle"][0]["total"]
            base["totales"]["total"] = base["detalle"][0]["total"]

        return base

    # ------------------ Guardado ------------------
    def save_json_for(self, source_path: Path, data: Dict) -> Path:
        rel = source_path.relative_to(self.extract_root)
        out_json = (self.out_json_root / rel).with_suffix(".json")
        out_json.parent.mkdir(parents=True, exist_ok=True)
        import json
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

            # Mapa por documento (document_id -> data)
            doc_map: Dict[str, Dict] = {}
            src_map: Dict[str, Path] = {}
            mode_map: Dict[str, str] = {}

            # ---------- 1) XMLs primero ----------
            xmls = [p for p in extracted_dir.rglob("*.xml") if p.is_file()]
            for xml in xmls:
                data = self.parse_invoice_from_ubl(xml)
                if not data:
                    continue

                doc_id = (
                        (data.get("document") or {}).get("document_id")
                        or f"XML::{xml.name}"
                )

                if doc_id in doc_map:
                    # Si hay duplicado, mergeamos por seguridad
                    doc_map[doc_id] = {
                        **doc_map[doc_id],
                        **_merge_dict(doc_map[doc_id], data),
                    }
                    mode_map[doc_id] = "XML"  # sigue siendo origen XML
                else:
                    doc_map[doc_id] = data
                    src_map[doc_id] = xml
                    mode_map[doc_id] = "XML"

            # ---------- 2) PDFs (enriquecer o crear si no hay XML) ----------
            pdfs = [p for p in extracted_dir.rglob("*.pdf") if p.is_file()]
            for pdf in pdfs:
                try:
                    text = self.pdf_to_text(pdf)
                    self.save_txt_for(pdf, text)
                    pdf_data = self.parse_from_text(text)
                except Exception as e:
                    self.log.error(f"PDF/OCR fail {pdf.name}: {e}")
                    res.errors += 1
                    continue

                doc_id_pdf = (
                        (pdf_data.get("document") or {}).get("document_id")
                        or f"PDF::{pdf.name}"
                )

                if doc_id_pdf in doc_map:
                    # Merge XML (primario) + PDF (secundario)
                    xml_data = doc_map[doc_id_pdf]
                    merged = {}

                    # Campos principales
                    merged["document"] = _merge_dict(
                        xml_data.get("document"), pdf_data.get("document")
                    )
                    merged["cufe"] = _merge_scalar(
                        xml_data.get("cufe"), pdf_data.get("cufe")
                    )
                    merged["cliente"] = _merge_dict(
                        xml_data.get("cliente"), pdf_data.get("cliente")
                    )
                    merged["emisor"] = _merge_dict(
                        xml_data.get("emisor"), pdf_data.get("emisor")
                    )
                    merged["fechas"] = _merge_dict(
                        xml_data.get("fechas"), pdf_data.get("fechas")
                    )
                    merged["pago"] = _merge_dict(
                        xml_data.get("pago"), pdf_data.get("pago")
                    )
                    merged["moneda"] = (
                            xml_data.get("moneda") or pdf_data.get("moneda") or "COP"
                    )
                    merged["totales"] = _merge_dict(
                        xml_data.get("totales"), pdf_data.get("totales")
                    )

                    # Detalle: si XML ya tiene detalle, lo mantenemos;
                    # si no, usamos el del PDF.
                    detalle_xml = xml_data.get("detalle") or []
                    detalle_pdf = pdf_data.get("detalle") or []
                    merged["detalle"] = detalle_xml or detalle_pdf

                    # Bloques nuevos
                    merged["tributario"] = _merge_dict(
                        xml_data.get("tributario"), pdf_data.get("tributario")
                    )
                    merged["dian"] = _merge_dict(
                        xml_data.get("dian"), pdf_data.get("dian")
                    )

                    doc_map[doc_id_pdf] = merged
                    mode_map[doc_id_pdf] = "XML+PDF"
                    # mantenemos el source original (XML) para el path de salida
                else:
                    # No había XML: este documento nace desde PDF
                    doc_map[doc_id_pdf] = pdf_data
                    src_map[doc_id_pdf] = pdf
                    mode_map[doc_id_pdf] = "PDF"

            # ---------- 3) Guardar todos los documentos consolidados del ZIP ----------
            for doc_id, data in doc_map.items():
                source = src_map.get(doc_id)
                if source is None:
                    # Fallback: si por alguna razón no hay source, usamos el propio zip
                    source = extracted_dir

                out_json = self.save_json_for(source, data)
                res.docs_json += 1
                index_rows.append(
                    {
                        "zip": zip_path.name,
                        "source": str(source),
                        "json": str(out_json),
                        "mode": mode_map.get(doc_id, "XML"),
                        "document_id": doc_id or "",
                        "serie": (data.get("document") or {}).get("serie") or "",
                    }
                )

        # ---------- 4) Índice CSV ----------
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
