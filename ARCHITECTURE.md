# AIVO — Arquitectura del Sistema RPA

## Descripcion General

**AIVO** es un sistema de Automatizacion Robotica de Procesos (RPA) para Windows que descarga facturas electronicas desde Outlook, las procesa y las ingresa automaticamente en el sistema **SAFIX/XENCO**. El sistema relaciona cada factura con un vehiculo (por placa) y su correspondiente centro de costos e interface contable.

---

## Estructura de Directorios

```
RPA_RENTAN_SAFIX/
├── main.py                              # Punto de entrada (lanza UI Flet)
├── requirements.txt                     # Dependencias Python
├── .env                                 # Variables de entorno (credenciales, rutas)
├── AIVO.spec                            # Configuracion PyInstaller
├── scripts/
│   └── build_exe.ps1                    # Script PowerShell para compilar EXE
│
├── src/
│   ├── config.py                        # Dataclass Settings (carga .env)
│   │
│   ├── core/                            # Infraestructura base
│   │   ├── com_init.py                  # Context manager COM (CoInitialize)
│   │   ├── logging_config.py            # Configuracion de logs (archivo + consola)
│   │   ├── progress_store.py            # Estado persistente entre corridas
│   │   ├── run_tracking.py              # Logs Excel por corrida
│   │   └── cleanup.py                   # Limpieza de archivos temporales
│   │
│   ├── outlook/                         # Integracion con Outlook
│   │   ├── client.py                    # Wrapper bajo nivel COM API
│   │   ├── service.py                   # Servicio alto nivel (descargar adjuntos)
│   │   ├── models.py                    # Clases de datos (MailSummary, SaveResult)
│   │   └── report_mailer.py             # Envio de reportes por correo
│   │
│   ├── processing/                      # Procesamiento de facturas
│   │   ├── zip_invoice_extractor.py     # Extraccion ZIP → XML/PDF → JSON
│   │   ├── aggregate_json.py            # Agregacion JSON por document_id
│   │   └── doc_id.py                    # Parser de IDs (DEFL/DENC)
│   │
│   ├── workflows/                       # Orquestacion de procesos
│   │   ├── invoice_pipeline.py          # Pipeline principal completo
│   │   ├── download_attachments.py      # Flujo de descarga desde Outlook
│   │   └── safix_automation.py          # Motor de automatizacion SAFIX
│   │
│   └── ui/                             # Interfaces de usuario
│       ├── ui_flet.py                   # UI de escritorio (Flet)
│       └── overlay_status.py            # Overlay Tkinter de progreso en tiempo real
│
├── tests/                               # Pruebas unitarias/integracion
├── downloads/                           # ZIPs descargados de Outlook
├── logs/                                # Logs globales persistentes
└── output/                              # Salidas por corrida (YYYY-MM-DD_HHMMSS)
```

---

## Flujo Principal del Sistema

```
┌─────────────────────────────────────────────────────────────────────┐
│                          USUARIO                                     │
│   1. Selecciona "Centros de Costos Vehiculos.xlsx"                   │
│   2. (Opcional) Ingresa rango de fechas                              │
│   3. Hace clic en "Ejecutar AIVO"                                    │
└───────────────────────────┬─────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│                     src/ui/ui_flet.py                                │
│   - Valida Excel (nombre exacto + cabeceras)                         │
│   - Valida fechas (formato YYYY-MM-DD, rango valido)                 │
│   - Lanza run_pipeline() en hilo separado (daemon thread)            │
└───────────────────────────┬─────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│              src/workflows/invoice_pipeline.py                       │
│                      run_pipeline()                                  │
│                                                                     │
│   FASE 1: Descarga    → run_download()                              │
│   FASE 2: Extraccion  → ZipInvoiceExtractor.process_all()          │
│   FASE 3: Agregacion  → build_invoices_by_id()                     │
│   FASE 4: SAFIX       → run_safix_with_excel()                     │
│   FASE 5: Logs Excel  → RunTracker.save()                          │
│   FASE 6: Limpieza    → cleanup_output_dir_keep_pdf_xml()          │
│   FASE 7: Correo      → send_report_email() (si habilitado)        │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Diagrama de Componentes

```
┌──────────────────────────────────────────────────────────────────────────────────┐
│                              AIVO — Componentes                                   │
│                                                                                   │
│  ┌──────────┐     ┌────────────────┐     ┌──────────────────────────────────┐    │
│  │ main.py  │────▶│  ui_flet.py    │────▶│    invoice_pipeline.py           │    │
│  └──────────┘     │  (Flet UI)     │     │    (Orquestador principal)       │    │
│                   └────────────────┘     └──────┬───────────────────────────┘    │
│                                                 │                                │
│              ┌──────────────────────────────────┼───────────────────────┐        │
│              │                                  │                       │        │
│              ▼                                  ▼                       ▼        │
│   ┌─────────────────────┐     ┌──────────────────────┐    ┌──────────────────┐   │
│   │ download_attachments│     │ zip_invoice_extractor│    │ safix_automation │   │
│   │   (Fase 1)          │     │   (Fases 2-3)        │    │   (Fase 4)       │   │
│   └──────────┬──────────┘     └──────────┬───────────┘    └────────┬─────────┘   │
│              │                           │                         │             │
│              ▼                           ▼                         ▼             │
│   ┌──────────────────┐      ┌────────────────────┐    ┌────────────────────────┐ │
│   │  outlook/        │      │  downloads/*.zip   │    │ PyAutoGUI / PyWinAuto  │ │
│   │  client.py       │      │  output/extract/   │    │ (automatizacion UI)    │ │
│   │  service.py      │      │  output/json/      │    └────────────────────────┘ │
│   │  (COM Outlook)   │      └────────────────────┘                              │
│   └──────────────────┘                                                          │
│                                                                                   │
│   ┌──────────────────────────────────────────────────────────────────────────┐   │
│   │                     INFRAESTRUCTURA TRANSVERSAL                          │   │
│   │  config.py  │  progress_store.py  │  run_tracking.py  │  logging_config  │   │
│   └──────────────────────────────────────────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────────────────────────────┘
```

---

## FASE 1 — Descarga desde Outlook

```
download_attachments.py
    └── OutlookService.save_attachments()
            └── OutlookClient.iter_items(days_back, only_unread, date_from, date_to)
                    │
                    │  Filtra correos via COM API (Outlook Desktop)
                    │  Filtrado Python-side para evitar problemas de formato regional
                    │
                    ▼
            Para cada correo con adjuntos:
                ├── Subcarpeta: downloads/YYYY-MM-DD/
                ├── Guarda adjunto: {nombre}__{sha256[:16]}.zip
                ├── Actualiza _manifest.json: { "zip_stem": ["entry_id1", ...] }
                └── Restaura UnRead=True si mark_as_read=False
                    (Outlook puede auto-marcar al acceder adjuntos via COM)
```

**Archivo de manifiesto** (`downloads/_manifest.json`):
```json
{
  "ZIP_NOMBRE__sha256": ["OUTLOOK_ENTRY_ID_1"],
  "OTRO_ZIP__sha256":   ["OUTLOOK_ENTRY_ID_2", "OUTLOOK_ENTRY_ID_3"]
}
```
> El manifiesto es clave para saber a que correo de Outlook corresponde cada ZIP, permitiendo marcar como leido despues de procesar en SAFIX.

---

## FASE 2 — Extraccion y Parsing de Facturas

```
ZipInvoiceExtractor.process_all()
    │
    ├── Para cada ZIP en downloads/:
    │       ├── Extrae en output/extract/{zip_stem}/
    │       │
    │       ├── Busca XML (formato UBL DIAN):
    │       │       ├── parse_invoice_from_ubl()
    │       │       │       ├── Detecta AttachedDocument (XML embebido)
    │       │       │       ├── Extrae: document_id, CUFE, NIT, fechas, totales, lineas
    │       │       │       └── _to_money_int(): maneja formato ES (1.000,00) y US (1,000.00)
    │       │       └── Guarda: output/json/{document_id}.json
    │       │
    │       └── Si no hay XML → busca PDF:
    │               ├── pdf_to_text() via pdfminer.six
    │               ├── Fallback OCR: pdf2image + pytesseract
    │               ├── parse_from_text(): regex para doc_id, montos
    │               └── Guarda: output/json/{document_id}.json
    │
    └── Genera index.csv con todos los JSONs procesados
```

**Schema JSON de factura**:
```json
{
  "document": { "serie": "DEFL", "numero": "12345678", "document_id": "DEFL12345678" },
  "cufe": "UUID del CUFE",
  "emisor": { "nombre": "Proveedor S.A.", "nit": "900000001" },
  "cliente": { "nombre": "RENTAN S.A.S", "nit": "860000001" },
  "fechas": { "venta": "2026-03-18T14:00:00", "vencimiento": "2026-04-17" },
  "totales": { "subtotal": 19100, "total": 19100 },
  "detalle": [{ "descripcion": "Servicio vehiculo ABC123", "cantidad": 1, "total": 19100 }],
  "_source_path": "C:/...output/extract/ZIP_STEM/factura.xml"
}
```

---

## FASE 3 — Agregacion por ID

```
build_invoices_by_id()
    │
    ├── Lee todos los JSONs de output/json/
    ├── Agrupa por document_id (evita duplicados)
    └── Genera: output/json/all_invoices_by_id.json
```

---

## FASE 4 — Automatizacion SAFIX

```
run_safix_with_excel(excel_path, aggregated_json_path)
    │
    ├── Carga catalogo de placas desde Excel:
    │       { "ABC123": { "interface": "INT01", "centro_costos": "CC001", "doble_cc": "" } }
    │
    ├── Carga facturas desde all_invoices_by_id.json
    │
    ├── SafixAutomator.bootstrap() — Lanza SAFIX si no esta abierto
    │       └── os.startfile(settings.safix_shortcut)  # abre .jnlp
    │
    ├── Para cada factura (en orden):
    │       │
    │       ├── 1. Extrae: placa (de descripcion linea), total, doc_id
    │       │
    │       ├── 2. Busca placa en catalogo Excel:
    │       │       ├── Si NO encontrada → log placas_no_encontradas.xlsx → SKIP (continue)
    │       │       └── Si encontrada → obtiene interface y centro_costos
    │       │
    │       ├── 3. Si DOBLE CC = "X" → log doble_cc.xlsx → SKIP
    │       │
    │       ├── 4. process_invoice() — Automatizacion UI:
    │       │       ├── Verifica ventana SAFIX abierta (pygetwindow)
    │       │       ├── Navega al formulario de causacion
    │       │       ├── Llena campos: NIT, doc_id, total, interface, centro_costos
    │       │       ├── Selecciona codigos desde dropdowns (imagen recognition o coordenadas)
    │       │       ├── Hace clic en guardar / submit
    │       │       └── Si error → screenshot → log FAIL → continua o detiene (segun config)
    │       │
    │       ├── 5. Marca correo como leido en Outlook:
    │       │       ├── _manifest_entry_ids_for_invoice(invoice, manifest, extract_root)
    │       │       │       └── Resuelve zip_stem desde _source_path → busca en manifest
    │       │       └── outlook_client.mark_as_read_by_id(entry_id) para cada ID
    │       │
    │       ├── 6. Actualiza progress_store (status=OK)
    │       │
    │       └── 7. Verifica pausa nocturna (si habilitada)
    │
    └── Al finalizar: genera overlay con resumen
```

---

## Pausa Nocturna (Night Pause)

Permite que el bot pause automaticamente en horas de la noche y se reactive:

```
_pause_if_night_window()
    │
    ├── _is_in_pause_window(now) → True si hora actual esta en [START, RESUME)
    │       └── Maneja wrap-around: 23:00 → 01:00 del dia siguiente
    │
    ├── Si en ventana de pausa:
    │       ├── Log "Pausando hasta HH:MM"
    │       ├── SafixAutomator.close_session_for_night_pause()
    │       │       └── Cierre graceful: icono puerta → ALT+A → ALT+S
    │       │
    │       └── Loop: sleep(check_sec) hasta que salga de ventana
    │               └── Cada iteracion: _next_resume_datetime()
    │
    └── Al salir de pausa:
            ├── SafixAutomator.bootstrap() → Relanza SAFIX
            └── Continua con siguiente factura
```

**Configuracion en .env**:
```ini
SAFIX_NIGHT_PAUSE_ENABLED=true
SAFIX_NIGHT_PAUSE_START=23:00
SAFIX_NIGHT_PAUSE_RESUME=01:00
SAFIX_NIGHT_PAUSE_CHECK_SEC=30
```

---

## Estado Persistente (Resume entre Corridas)

```
src/logs/resume_state.json
    {
      "invoices": {
        "DEFL12345678": {
          "status": "OK",
          "placa": "ABC123",
          "updated_at": "2026-03-18T14:30:00"
        },
        "DENC87654321": {
          "status": "FAIL",
          "error": "Ventana SAFIX no responde",
          "updated_at": "2026-03-18T15:00:00"
        }
      }
    }
```

- `OK` → Se omite en proximas corridas
- `FAIL` → Se reintenta en proxima corrida
- `MISSING` → Se reintenta cuando la placa sea agregada al Excel
- `DOBLE_CC` → No se procesa (marcado explicitamente en catalogo)

---

## Logs Generados

### Por Corrida (`output/YYYY-MM-DD_HHMMSS/logs/`)

| Archivo | Contenido |
|---|---|
| `descargados.xlsx` | ZIPs descargados: nombre, ruta, tamano, estado |
| `procesadas.xlsx` | Facturas procesadas: doc_id, placa, total, interface, CC, estado |
| `placas_no_encontradas.xlsx` | Facturas sin placa en catalogo |
| `doble_cc.xlsx` | Facturas omitidas por DOBLE CC |

### Globales Persistentes (`src/logs/`)

| Archivo | Contenido |
|---|---|
| `app_YYYY-MM-DD.log` | Log de texto de la aplicacion (rotado diariamente) |
| `resume_state.json` | Estado de cada factura (para resume entre corridas) |
| `facturas_descargadas.xlsx` | Historial acumulado de descargas |
| `facturas_procesadas_global.xlsx` | Historial acumulado de procesadas |

---

## Catalogo de Vehiculos (Excel)

Archivo: `Centros de Costos Vehiculos.xlsx`

| N VEHICULO | PLACA | UBICACION | CENTRO DE COSTOS | INTERFACE | DOBLE CC |
|---|---|---|---|---|---|
| 001 | ABC123 | BOGOTA | CC001 | INT01 | |
| 002 | XYZ789 | MEDELLIN | CC002 | INT02 | X |

- **DOBLE CC = "X"**: la factura se omite (va a otro sistema)
- **PLACA**: normalizada a mayusculas sin separadores (`ABC-123` → `ABC123`)
- El archivo debe llamarse exactamente `Centros de Costos Vehiculos.xlsx`

---

## Deduplicacion de ZIPs

Los archivos ZIP se guardan con hash SHA-256 en el nombre para evitar duplicados:

```
{nombre_original}__{sha256[:16]}.zip
```

Si el archivo destino ya existe, se omite (el ZIP fue procesado antes). Esto permite re-ejecutar el pipeline sin reprocessar ZIPs ya descargados.

---

## Clases Principales

### `OutlookClient` (`src/outlook/client.py`)

```python
class OutlookClient:
    def find_store_by_display(account: str) → Store
    def get_folder(store, path: list[str]) → Folder
    def iter_items(folder, days_back, only_unread, date_from, date_to) → Iterator[Item]
    def mark_as_read(item) → None
    def mark_as_read_by_id(entry_id: str) → None   # Para marcar post-SAFIX
    def attachments(item) → list[Attachment]
    def move_to(item, folder) → None
```

### `SafixAutomator` (`src/workflows/safix_automation.py`)

```python
class SafixAutomator:
    def bootstrap() → bool               # Lanza SAFIX si no esta abierto
    def is_main_window_open() → bool      # Verifica ventana activa
    def process_invoice(invoice_data) → InvoiceResult
    def refocus_main() → None
    def write_text_safe(text) → None      # Ctrl+A, backspace, write (robusto)
    def capture_error_snapshot() → Path   # Screenshot en error
    def close_session_for_night_pause()   # Cierre graceful
```

### `ProgressStore` (`src/core/progress_store.py`)

```python
class ProgressStore:
    def is_processed(document_id: str) → bool
    def mark(document_id: str, status: str, **meta) → None
    def load() → dict
    def save() → None
```

### `RunTracker` (`src/core/run_tracking.py`)

```python
class RunTracker:
    def add_download(original_name, saved_path, size_bytes, status, error) → None
    def add_processed(document_id, placa, total, interface, cc, status, error) → None
    def add_missing_plate(...) → dict
    def append_missing_plate_row(row: dict) → None
    def save() → dict[str, Path]   # Retorna rutas de los Excel generados
```

---

## Manejo de Errores

| Escenario | Comportamiento |
|---|---|
| Placa no en catalogo | Log a `placas_no_encontradas.xlsx`, skip, status=MISSING (reintentable) |
| DOBLE CC marcado | Log a `doble_cc.xlsx`, skip, no se procesa en SAFIX |
| Error en SAFIX | Screenshot, log FAIL, continua con siguiente (si `SAFIX_STOP_ON_ERROR=false`) |
| Ventana SAFIX cerrada | Relanza automaticamente si `safix_recover_session_on_error=true` |
| Email ya leido | Restaura `UnRead=True` si `mark_as_read=False` (Outlook auto-marca al leer adjuntos) |
| ZIP ya descargado | SHA-256 dedup: se salta si archivo destino ya existe |
| Factura ya procesada | `progress_store` skipea en proximas corridas (status=OK) |
| SAFIX cuelgue (N facturas) | `SAFIX_SOFT_RESET_EVERY=N` reinicia sesion cada N facturas |

---

## Compilacion a EXE (PyInstaller)

```powershell
# scripts/build_exe.ps1
pyinstaller AIVO.spec --clean
```

El `AIVO.spec` incluye:
- Todos los archivos de `src/`
- Assets (iconos, imagenes de referencia para automation)
- Dependencias COM (pywin32)
- Tesseract (OCR) si aplica
- Configuracion de ventana sin consola (`console=False`)

**Salida**: `dist/AIVO.exe`

---

## Configuracion (.env)

```ini
# Outlook
OUTLOOK_ACCOUNT=correo@empresa.com
OUTLOOK_FOLDER=Facturas
DOWNLOAD_DIR=C:/RPA/downloads
PROCESSED_FOLDER=Procesados

# SAFIX
SAFIX_SHORTCUT=C:/path/to/safix.jnlp
SAFIX_WINDOW_TITLE=SAFIX.*
SAFIX_USER=usuario
SAFIX_PASS=contraseña
SAFIX_NIT=860000001

# Codigos
SAFIX_XOT_CODE=XOT01
SAFIX_GOT_CODE=GOT01
SAFIX_CAMPO_84=84
SAFIX_CAMPO_05=05

# Tiempos (segundos)
SAFIX_FORM_READY_WAIT=50.0
SAFIX_PYAUTO_PAUSE=1.1
SAFIX_WAIT_DEFAULT=1.5
SAFIX_WAIT_LONG=3.0

# Pausa nocturna
SAFIX_NIGHT_PAUSE_ENABLED=false
SAFIX_NIGHT_PAUSE_START=23:00
SAFIX_NIGHT_PAUSE_RESUME=01:00

# Reportes
REPORT_EMAIL_ENABLED=false
REPORT_EMAIL_TO=jefe@empresa.com
```

---

## Flujo de Datos Completo

```
Outlook (correos)
    │
    │  [COM API: win32com]
    ▼
downloads/
    ├── 2026-03-18/factura_ABC__a1b2c3d4.zip
    └── _manifest.json  {"factura_ABC__a1b2c3d4": ["ENTRY_ID_123"]}
    │
    │  [ZipInvoiceExtractor]
    ▼
output/2026-03-18_093924/
    ├── extract/
    │   └── factura_ABC__a1b2c3d4/
    │       └── factura.xml
    ├── json/
    │   ├── DEFL12345678.json     ← Invoice parseada
    │   └── all_invoices_by_id.json
    │
    │  [run_safix_with_excel]
    ▼
Catalogo Excel (PLACA → CC + Interface)
    │
    │  [PyAutoGUI / PyWinAuto]
    ▼
SAFIX/XENCO (causacion de facturas)
    │
    │  [mark_as_read_by_id]
    ▼
Outlook (correo marcado como leido)
    │
    ▼
output/logs/
    ├── descargados.xlsx
    ├── procesadas.xlsx
    ├── placas_no_encontradas.xlsx
    └── doble_cc.xlsx
```

---

## Dependencias Principales

| Libreria | Version | Uso |
|---|---|---|
| `pywin32` | 311 | COM API (Outlook, Windows) |
| `pywinauto` | 0.6.9 | Automatizacion ventanas Windows |
| `PyAutoGUI` | 0.9.54 | Control teclado/mouse |
| `flet` | 0.28.3 | UI de escritorio multiplataforma |
| `openpyxl` | 3.1.5 | Lectura/escritura Excel |
| `pdfminer.six` | — | Extraccion texto de PDFs |
| `pdf2image` | — | PDF a imagen (para OCR) |
| `pytesseract` | — | OCR (requiere Tesseract instalado) |
| `python-dotenv` | 1.2.1 | Carga de `.env` |
| `Pillow` | 12.0.0 | Procesamiento de imagenes |
