# RPA Outlook + SAFIX (Windows / COM)

Automatiza la descarga de adjuntos desde Outlook Desktop, extrae y procesa facturas (XML/PDF), y carga datos en SAFIX/XENCO por automatizacion de UI.

Tambien incluye:
- reanudacion de proceso (resume) entre ejecuciones,
- logs en Excel en tiempo real,
- envio opcional de reportes por correo al finalizar.

## Requisitos
- Windows 10/11
- Outlook Desktop con sesion iniciada
- SAFIX/XENCO instalado y acceso al `.jnlp`
- Python 3.9+
- Misma arquitectura entre Python y Outlook (x64/x86)

## Instalacion
```bash
pip install -r requirements.txt
```

## Ejecucion
Entrada principal:
```bash
python main.py
```

Alternativa:
```bash
python -m src.ui.ui_flet
```

## Flujo actual
1. Descarga adjuntos ZIP desde Outlook.
2. Extrae ZIPs y genera JSON individuales por factura.
3. Ejecuta SAFIX con esas facturas.
4. Marca correo como leido solo cuando la factura queda `OK` en SAFIX.
5. Genera reportes y estado de reanudacion.

## Salidas por ejecucion
Se crea carpeta con fecha/hora:
- `output/YYYY-MM-DD_HHMMSS/`

Archivos principales:
- `output/.../json/index.csv`
- `output/.../logs/descargados.xlsx`
- `output/.../logs/procesadas.xlsx`
- `output/.../logs/placas_no_encontradas.xlsx`
- `output/.../logs/doble_cc.xlsx`

## Estado global de reanudacion (persistente)
Se guarda en `src/logs/` (no depende de la carpeta de fecha):
- `src/logs/resume_state.json`
- `src/logs/facturas_descargadas.xlsx`
- `src/logs/facturas_procesadas_global.xlsx`

Uso:
- Si una factura ya esta en estado `OK`, no se reprocesa en la siguiente corrida.
- Si una corrida se interrumpe, la siguiente retoma pendientes.

## Logging unificado
Todo lo impreso en consola se guarda en:
- `src/logs/app_YYYY-MM-DD.log`

## Excel de placas (UI)
Nombre esperado:
- `Centros de Costos Vehiculos.xlsx`

Columnas requeridas:
- `N° VEHICULO`
- `PLACA`
- `UBICACIÓN`
- `CENTRO DE COSTOS`
- `INTERFACE`
- `DOBLE CC`

Regla `DOBLE CC`:
- Si la celda contiene `X`, la placa se omite en SAFIX.
- Se registra en:
  - `placas_no_encontradas.xlsx`
  - `doble_cc.xlsx`

## Variables `.env` relevantes
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
```

Robustez:
```env
SAFIX_STOP_ON_ERROR=false
SAFIX_SOFT_RESET_EVERY=0
SAFIX_LOG_EVERY=1
SAFIX_BREATH_SEC=0.4
SAFIX_ERROR_DIR=./output/errors
SAFIX_PREFLIGHT_ICONS=false
SAFIX_SCREENSHOT_ON_ERROR=true
```

Correo de reportes (opcional):
```env
REPORT_EMAIL_ENABLED=false
REPORT_EMAIL_TO=
REPORT_EMAIL_CC=
REPORT_EMAIL_SUBJECT=Reporte automatizacion SAFIX
REPORT_EMAIL_BODY=Adjunto reportes de la ejecucion automatizada.
```

## Envio de correo al finalizar
Si `REPORT_EMAIL_ENABLED=true`, el pipeline envia correo por Outlook COM con adjuntos disponibles:
- `descargados.xlsx`
- `procesadas.xlsx`
- `placas_no_encontradas.xlsx`
- `doble_cc.xlsx`
- `index.csv`
- `resume_state.json`
- `facturas_descargadas.xlsx`
- `facturas_procesadas_global.xlsx`

## Pruebas
```bash
python -m pytest -q
```

Pruebas agregadas de regresion:
- `tests/test_pipeline_regressions.py`
- `tests/test_progress_store.py`

## Notas operativas
- En la UI se usan: `Solo no leidos` y `Marcar como leidos`.
- El marcado como leido es posterior al `OK` en SAFIX.
- Si no hay facturas, no se inicia SAFIX y se muestra mensaje.
