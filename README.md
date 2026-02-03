# RPA Outlook + SAFIX (COM-based)

Automatiza Microsoft Outlook Desktop (Windows) vía COM con pywin32, procesa facturas (XML/PDF/OCR), consolida por `document_id` y carga en SAFIX/XENCO mediante automatización visual.

Incluye UI en Flet para ejecutar el pipeline completo y logs en Excel en tiempo real.

**Características principales**
- Acceso directo a Outlook Desktop usando COM (pywin32)
- Descarga controlada de adjuntos con filtros por fecha y no leídos
- Extracción automática de ZIPs
- Procesamiento de XML y PDF (OCR cuando aplica)
- Consolidación por factura (`document_id`)
- Automatización visual de SAFIX (PyAutoGUI + PyWinAuto)
- Overlay de estado en tiempo real
- Logs Excel en tiempo real de procesadas y placas no encontradas
- Captura de pantalla automática en errores (opcional)
- Configuración centralizada mediante `.env`
- Interfaz gráfica (Flet)

**Estructura del proyecto**
```
.
├─ src/
│  ├─ config.py
│  ├─ assets/
│  ├─ core/
│  ├─ outlook/
│  ├─ processing/
│  ├─ workflows/
│  └─ ui/
├─ output/
├─ logs/
├─ main.py
├─ .env
├─ README.md
└─ requirements.txt
```

**Requerimientos**
- Windows 10/11
- Outlook Desktop instalado y con sesión iniciada
- SAFIX/XENCO instalado
- Acceso al archivo `.jnlp` de SAFIX
- Python 3.9+
- Coincidencia de arquitectura (x64/x86) entre Python y Outlook

**Instalación**
```bash
pip install -r requirements.txt
```

**Ejecución**
```bash
python main.py
```

También funciona:
```bash
python -m src.ui.ui_flet
```

**Flujo general**
1. Descargar adjuntos desde Outlook (ZIP).
2. Extraer ZIPs.
3. Procesar XML/PDF y generar JSON por factura.
4. Automatizar carga en SAFIX usando el Excel de placas.
6. Registrar logs en Excel y limpiar archivos temporales.

**Archivos de salida**
- JSONs: `output/json/`
- Agregado: no se genera (se usan JSON individuales).
- PDF/XML: `output/extract/`
- Logs: `output/logs/procesadas.xlsx`
- Placas no encontradas: `output/logs/placas_no_encontradas.xlsx`
- Errores (screenshots): `output/errors/`

**Configuración (.env)**
Ejemplo base:
```env
OUTLOOK_ACCOUNT=CORREO
OUTLOOK_FOLDER=NOMBRE_DE_CARPETA_EN_CORREO
DOWNLOAD_DIR=./downloads
LOG_DIR=./src/logs
OUTPUT_DIR=./output

SAFIX_SHORTCUT=C:/ruta/a/safix.jnlp
SAFIX_WINDOW_TITLE=.*XENCO - Administracion del Sistema.*
SAFIX_USER=usuario
SAFIX_PASS=clave
SAFIX_NIT=123456789
SAFIX_XOT_CODE=XOT05
SAFIX_GOT_CODE=GOT
SAFIX_CCAMPO_84=84
SAFIX_CAMPO_05=05
SAFIX_OBL_CODE=OBL_EXCLU
SAFIX_OBL2_CODE=OBL_ANTCON
SAFIX_TESORERIA_ICON=src/assets/tesoreria.png
SAFIX_VALORES_ICON=src/assets/valores.png
SAFIX_Z_ICON=src/assets/Z.png
SAFIX_ENGRANES_ICON=src/assets/engranes.png
```

**Logs unificados**
- Todo lo que se imprime en consola se guarda en un único archivo por fecha:
  `src/logs/app_YYYY-MM-DD.log`

**Output con fecha**
- Cada ejecución genera una carpeta con timestamp:
  `output/YYYY-MM-DD_HHMMSS/`

Opciones de robustez SAFIX:
```env
SAFIX_STOP_ON_ERROR=true
SAFIX_SOFT_RESET_EVERY=25
SAFIX_LOG_EVERY=1
SAFIX_BREATH_SEC=0.4
SAFIX_ERROR_DIR=./output/errors
SAFIX_PREFLIGHT_ICONS=false
SAFIX_SCREENSHOT_ON_ERROR=true
```

**Notas de uso**
- La UI permite seleccionar rango de fechas, solo no leídos y marcar como leídos.
- El Excel debe llamarse exactamente `Centros de Costos Vehiculos.xlsx`.
- Cabeceras requeridas: `N° VEHICULO | PLACA | UBICACIÓN | CENTRO DE COSTOS | INTERFACE`

**Pruebas**
```bash
pytest -q
```

**Seguridad**
- No se almacenan credenciales en el código.
- Outlook COM accede al perfil del usuario autenticado.

**Limitaciones**
- Solo Windows con Outlook Desktop.
- Requiere sesión activa en Outlook.
- No funciona como servicio sin usuario logueado.

**Contacto y mantenimiento**
Autor: Michaen Stebin Rangel Giraldo  
Rol: Arquitecto de Soluciones/Software/RPA  
Proyecto: Rentan E.I.C.E – Inter-Telco S.A.S
