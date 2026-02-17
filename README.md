# RPA Outlook + SAFIX (Windows / COM)

Automatiza descarga de adjuntos desde Outlook, extracción de facturas (XML/PDF) y carga en SAFIX/XENCO por UI automation.

## Requisitos
- Windows 10/11
- Outlook Desktop con sesión iniciada
- SAFIX/XENCO disponible por `.jnlp`
- Python 3.9+
- Mismo tipo de arquitectura entre Python y Outlook (x64/x86)

## Ejecución
```bash
python main.py
```

## UI actual (simplificada)
La interfaz en `src/ui/ui_flet.py` quedó reducida a:
- Botón para cargar Excel
- Botón `Ejecutar AIVO`
- Panel de estado/error multilínea

Parámetros fijos en UI (sin controles visibles):
- `only_unread=True`
- `mark_as_read=True`
- `days_back=0`

## Flujo
1. Descarga ZIPs desde Outlook.
2. Extrae y genera JSON por factura.
3. Ejecuta SAFIX por factura.
4. Marca correo como leído cuando la factura queda `OK`.
5. Genera reportes.

## Salidas
Por ejecución:
- `output/YYYY-MM-DD_HHMMSS/`

Reportes típicos:
- `output/.../json/index.csv`
- `output/.../logs/descargados.xlsx`
- `output/.../logs/procesadas.xlsx`
- `output/.../logs/placas_no_encontradas.xlsx`
- `output/.../logs/doble_cc.xlsx`

Estado global (persistente):
- `src/logs/resume_state.json`
- `src/logs/facturas_descargadas.xlsx`
- `src/logs/facturas_procesadas_global.xlsx`

## Logging
- Log unificado diario: `src/logs/app_YYYY-MM-DD.log`

## Excel requerido
Archivo:
- `Centros de Costos Vehiculos.xlsx`

Cabeceras requeridas:
- `N° VEHICULO`
- `PLACA`
- `UBICACIÓN`
- `CENTRO DE COSTOS`
- `INTERFACE`
- `DOBLE CC`

Regla `DOBLE CC`:
- Si tiene `X`, la placa se omite y se reporta.

## Variables `.env` clave
Base:
```env
OUTLOOK_ACCOUNT=
OUTLOOK_FOLDER=
DOWNLOAD_DIR=./downloads
LOG_DIR=./src/logs
OUTPUT_DIR=./output
ENABLE_CLEANUP=false
```

SAFIX:
```env
SAFIX_SHORTCUT=
SAFIX_WINDOW_TITLE=.*XENCO - Administracion del Sistema.*
SAFIX_USER=
SAFIX_PASS=
SAFIX_NIT=
SAFIX_XOT_CODE=XOT05
SAFIX_GOT_CODE=GOT
SAFIX_CAMPO_84=84
SAFIX_CAMPO_05=05
SAFIX_OBL_CODE=OBL_EXCLU
SAFIX_OBL2_CODE=OBL_ANTCON
SAFIX_TESORERIA_ICON=src/assets/tesoreria.png
SAFIX_VALORES_ICON=src/assets/valores.png
SAFIX_Z_ICON=src/assets/Z.png
SAFIX_ENGRANES_ICON=src/assets/engranes.png
SAFIX_DOOR_ICON=src/assets/salir.png
SAFIX_CONN_ICON=src/assets/conn.png
```

Mitigaciones:
```env
SAFIX_STOP_ON_ERROR=false
SAFIX_SOFT_RESET_EVERY=0
SAFIX_LOG_EVERY=1
SAFIX_BREATH_SEC=0.4
SAFIX_ERROR_DIR=./output/errors
SAFIX_PREFLIGHT_ICONS=false
SAFIX_SCREENSHOT_ON_ERROR=true
SAFIX_NIGHT_PAUSE_ENABLED=false
SAFIX_NIGHT_PAUSE_START=23:00
SAFIX_NIGHT_PAUSE_RESUME=01:00
SAFIX_NIGHT_PAUSE_CHECK_SEC=30
SAFIX_RECOVER_SESSION_ON_ERROR=true
```

Pausa nocturna:
- Si `SAFIX_NIGHT_PAUSE_ENABLED=true`, el bot pausa entre facturas dentro de la ventana `START -> RESUME`.
- Durante pausa nocturna cierra sesi?n de forma controlada: click `door_icon` -> `ALT+A` -> `ALT+S` -> click `conn_icon`.
- Luego espera hasta la hora de reanudaci?n y vuelve a abrir sesi?n.
- Al reanudar contin?a pendientes (no reprocesa `OK` por el estado persistente).

Correo (opcional):
```env
REPORT_EMAIL_ENABLED=false
REPORT_EMAIL_TO=
REPORT_EMAIL_CC=
REPORT_EMAIL_SUBJECT=Reporte automatizacion SAFIX
REPORT_EMAIL_BODY=Adjunto reportes de la ejecucion automatizada.
```

## Pruebas
```bash
python -m pytest -q
```

