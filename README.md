# Outlook Automation Bot (COM-based)

Automatiza Microsoft Outlook (versión de escritorio para Windows) mediante la interfaz COM (Component Object Model) usando la librería pywin32.
Permite listar correos, descargar adjuntos, marcar mensajes como leídos y moverlos a carpetas específicas, todo configurable desde variables de entorno.

Ideal para procesos corporativos donde Outlook es la herramienta principal de correo y se desea extraer o procesar información de manera automática sin depender de interfaces web o Selenium.

## Características principales

- Acceso directo a Outlook Desktop usando COM (sin clics ni UI automation)
- Descarga de adjuntos desde carpetas específicas (ej. FLYPASS, FACTURAS, etc.)
- Filtrado por días y estado de lectura
- Opción para marcar correos como leídos
- Opción para mover correos procesados a otra carpeta
- Logs rotativos detallados
- Configuración sencilla mediante .env
- Pruebas unitarias con mocks (no requiere Outlook real)
- Arquitectura modular y escalable

## Estructura del proyecto

```
outlook-bot/
├─ src/
│  ├─ core/
│  │  ├─ com_init.py              # Contexto COM seguro
│  │  └─ logging_config.py        # Configuración centralizada de logs
│  ├─ config.py                   # Carga de variables .env
│  ├─ outlook/
│  │  ├─ models.py                # Modelos de datos (MailSummary, SaveResult)
│  │  ├─ client.py                # Cliente COM (bajo nivel)
│  │  └─ service.py               # Lógica de negocio (guardar adjuntos, mover, etc.)
│  └─ workflows/
│     └─ download_attachments.py  # Flujo principal de descarga
├─ tests/
│  ├─ test_client.py              # Pruebas de estructura del cliente
│  └─ test_workflow.py            # Pruebas del flujo con mocks
├─ .env                           # Configuración de entorno
├─ main.py                        # CLI principal
├─ requirements.txt               # Dependencias
└─ pytest.ini                     # Config opcional para pytest
```

## Requerimientos

### Sistema operativo
- Windows 10 / 11
- Outlook Desktop instalado y con sesión iniciada

### Entorno Python
- Python 3.9 o superior
- Coincidencia de arquitectura (x64/x86) entre Python y Outlook

### Dependencias
```
pip install -r requirements.txt
```

Contenido de requirements.txt:
```
pywin32>=306
python-dateutil>=2.9.0
python-dotenv>=1.0.1
pytest>=8.2.0
```

## Configuración inicial (.env)

Ejemplo de archivo .env:

```
OUTLOOK_ACCOUNT=CORREO
OUTLOOK_FOLDER=NOMBRE_DE_CARPETA_EN_CORREO
DOWNLOAD_DIR=./downloads
ONLY_UNREAD=false
DAYS_BACK=30
MARK_AS_READ=true
MOVE_TO_PROCESSED=false
PROCESSED_FOLDER=Procesados
LOG_DIR=./logs
```

Todos los valores son opcionales. Si no se especifican, el código usa los valores predeterminados mostrados arriba.

## Cómo funciona

### 1. Inicialización COM
- Se crea un contexto seguro (CoInitialize/CoUninitialize) con el helper com_initialized().
- Se lanza Outlook internamente (si no está abierto) y se accede al namespace MAPI.

### 2. Cliente Outlook
- OutlookClient es una capa delgada sobre los objetos COM reales de Outlook:
  - find_store_by_display() encuentra el buzón (store).
  - get_folder() abre carpetas anidadas.
  - iter_items() recorre mensajes con filtros (days_back, unread).
  - attachments() devuelve los objetos Attachment.

### 3. Servicio Outlook
- OutlookService combina el cliente con la lógica de negocio:
  - Guarda adjuntos en carpetas organizadas por fecha.
  - Evita duplicados (usa SHA-256 para diferenciar archivos).
  - Marca como leído el correo procesado.
  - Mueve los correos si MOVE_TO_PROCESSED=true.

### 4. Logging
- Configurado por logging_config.py con:
  - Consola y archivo rotativo (logs/outlook_bot.log).
  - Formato:  
    2025-10-20 10:32:15 | INFO | outlook_bot | Processing 'Factura 2025...'

### 5. Workflow
- src/workflows/download_attachments.py ejecuta el flujo completo:
  1. Carga configuración del .env.
  2. Inicializa COM.
  3. Abre la carpeta definida (ej. FLYPASS).
  4. Descarga todos los adjuntos de los correos filtrados.
  5. Registra los resultados y finaliza.

## Ejecución

Ejecuta el comando principal:

```
python main.py --download
```

Salida esperada:
```
Processed: 42 | Attachments saved: 128
```

Los adjuntos se guardarán en ./downloads/YYYY-MM-DD/

## Pruebas

Puedes ejecutar las pruebas unitarias sin Outlook real (mockeadas):

```
pytest -q
```

Qué validan:
- La estructura de OutlookClient.
- El flujo completo de descarga (OutlookService.save_attachments) usando mails simulados.
- La creación de archivos y conteo correcto de adjuntos.

## Estructura de logs

Ubicación por defecto: ./logs/outlook_bot.log

Ejemplo de registro:
```
2025-10-20 09:55:32 | INFO | outlook_bot | Using store: "CORREO" | folder: "NOMBRE_DE_CARPETA_EN_CORREO"
2025-10-20 09:55:32 | INFO | outlook_bot | Processing: 'Factura Electrónica 2025-10-18'
2025-10-20 09:55:33 | INFO | outlook_bot | Processed=1 | Saved=2
```

Los logs rotan automáticamente cuando superan 2 MB (con 5 backups).

## Extensiones posibles

- Listar carpetas disponibles
  python main.py --folders
- Listar correos
  python main.py --list
- Integración con Pywinauto
  Añadir un módulo src/apps/<nombre_app> para interactuar con otra aplicación de escritorio.
- Integración con bases de datos
  Guardar los metadatos (asunto, remitente, hash del adjunto) en PostgreSQL o SQLite.
- Tareas programadas
  Ejecutar python main.py --download cada hora mediante el Programador de tareas de Windows.

## Seguridad

- El script no almacena contraseñas ni credenciales.
- Outlook COM accede directamente al perfil del usuario logueado.
- Si Outlook muestra un prompt de seguridad, ajusta:
  Archivo → Opciones → Centro de confianza → Seguridad del programa.

## Limitaciones

- Solo funciona en Windows con Outlook Desktop (no Outlook Web).
- Requiere sesión activa en Outlook.
- No se ejecuta como servicio en segundo plano sin usuario logueado.
- Las carpetas deben coincidir con el idioma de Outlook.

## Cómo trabaja internamente

1. COM Client (Outlook.Application)
   Pywin32 solicita a Windows el objeto registrado bajo Outlook.Application → Windows carga OUTLOOK.EXE si no está en ejecución.
2. Namespace MAPI
   GetNamespace("MAPI") expone todos los buzones y carpetas (Stores y Folders).
3. Recorrido y descarga
   El script accede a la carpeta configurada, recorre folder.Items, filtra por fecha y estado, y guarda los adjuntos mediante attachment.SaveAsFile(path).
4. Gestión de COM
   El contexto com_initialized() garantiza CoInitialize y CoUninitialize por hilo.
5. Logs
   Todos los pasos se registran con timestamp, nivel y módulo.

## Contacto y mantenimiento

Autor: Michaen Stebin Rangel Giraldo  
Rol: Arquitecto de Soluciones/Software/RPA  
Proyecto: Automatización Outlook – Inter-Telco S.A.S