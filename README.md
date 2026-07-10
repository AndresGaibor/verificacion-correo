# Verificación de Correos OWA

Herramienta Python para extraer información de contacto del directorio OWA de Madrid (correoweb.madrid.org/owa). Soporta dos métodos de extracción: **API REST** (recomendado) y **Playwright** (legacy).

## Características

- **API REST** — Extracción vía FindPeople + GetPersona (Exchange Web Services). Rápido, sin navegador.
- **GAL Scraper** — Extracción completa del directorio global (Global Address List) con paginación. Reanudable si la sesión expira.
- **Playwright** — Método legacy que interactúa con la UI de OWA.
- **Procesamiento incremental** — Solo procesa correos pendientes. Reanudable tras fallos de sesión.
- **GUI + CLI** — Interfaz gráfica tkinter y línea de comandos completa.
- **Sistema reanudable** — Detecta sesión expirada (HTTP 307) y guarda progreso.

## Instalación

```bash
git clone https://github.com/AndresGaibor/verificacion-correo.git
cd verificacion-correo
python3 -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -e .
pip install playwright openpyxl pyyaml
playwright install chromium
```

## Uso Rápido

```bash
# 1. Configurar sesión (login manual en navegador)
verificacion-correo setup

# 2. Procesar correos pendientes (vía Playwright)
verificacion-correo

# 3. Procesar correos vía API REST (más rápido)
verificacion-correo --api

# 4. Extraer todo el directorio GAL
verificacion-correo scrape-gallery

# 5. GUI
verificacion-correo-gui
```

## Comandos CLI

| Comando | Descripción |
|---------|-------------|
| `process` (default) | Procesa correos pendientes desde Excel |
| `setup` | Configura sesión del navegador (login manual) |
| `validate` | Valida configuración y preparación |
| `status` | Muestra estado de sesión y archivos |
| `scrape-gallery` | Extrae el directorio GAL completo vía API |

### Opciones de `process`

| Opción | Descripción |
|--------|-------------|
| `--excel-file PATH` | Archivo Excel específico |
| `--batch-size N` | Tamaño de lote (default: 10) |
| `--dry-run` | Vista previa sin procesar |
| `--force` | Forzar aunque la sesión sea inválida |
| `--keep-draft` | Mantener borrador abierto al terminar |

### Opciones de `scrape-gallery`

| Opción | Descripción |
|--------|-------------|
| `--output-dir DIR` | Directorio de salida (default: data/gal) |
| `--max-contacts N` | Máx. contactos (0 = ilimitado) |
| `--batch-size N` | Entradas por llamada API (default: 100) |
| `--delay SEC` | Segundos entre peticiones (default: 8) |
| `--force-restart` | Ignorar progreso guardado y empezar de cero |

### Aliases

```bash
verificacion-correo   # CLI
vcorreo               # Alias corto
verificacion-correo-gui  # GUI
```

## Métodos de Extracción

### API REST (Recomendado)

Usa los servicios EWS `FindPeople` + `GetPersona` con las cookies de sesión. Dos pasos:

1. **FindPeople** — Busca en el directorio por email, obtiene `PersonaId`
2. **GetPersona** — Obtiene detalles completos (teléfono, dirección, departamento, SIP, etc.)

Ventajas: sin navegador, más rápido, menos recursos.
Implementación: `src/verificacion_correo/core/api_extractor.py`

**Límite conocido**: La sesión OWA expira tras ~40-50 llamadas API (HTTP 307 → ADFS login). El sistema detecta esto, marca el error y los correos restantes quedan como pendientes para el siguiente ciclo.

### Playwright (Legacy)

Navega por la interfaz OWA: abre un nuevo mensaje, añade destinatarios, hace clic en tokens y extrae datos del popup de contacto.
Implementación: `src/verificacion_correo/core/browser.py`

### GAL Scraper

Extracción paginada del directorio global (sin filtro, `QueryString: None`). Guarda progreso en `gal_progress.json` para reanudar si la sesión expira. Exporta a JSON + CSV.
Implementación: `src/verificacion_correo/core/gal_scraper.py`

## Datos Extraídos

| Campo | Ejemplo |
|-------|---------|
| Email personal | nombre.apellido@madrid.org |
| Teléfono | 916704092 |
| Dirección | C/ AYUNTAMIENTO, 5 28791 RIVAS-VACIAMADRID |
| Departamento | OFICINA JUDICIAL MUNICIPAL |
| Compañía | ORGANOS JUDICIALES |
| Oficina | RIVAS-VACIAMADRID |
| SIP | sip:asp164@madrid.org |
| **Nombre** | ❌ Bloqueado por anti-scraping de OWA |

## Estructura del Proyecto

```
verificacion-correo/
├── src/verificacion_correo/
│   ├── __init__.py
│   ├── __main__.py              # python -m verificacion_correo → CLI
│   ├── core/
│   │   ├── api_extractor.py      # Extracción vía FindPeople + GetPersona (API REST)
│   │   ├── browser.py            # Automatización con Playwright
│   │   ├── config.py             # Gestión de configuración YAML
│   │   ├── excel.py              # Lectura/escritura Excel incremental
│   │   ├── extractor.py          # Extracción de contactos desde popups OWA
│   │   ├── first_run.py          # Configuración inicial automática
│   │   ├── gal_scraper.py        # Scraper del directorio GAL (paginated FindPeople)
│   │   └── session.py            # Gestión de sesiones del navegador
│   ├── cli/
│   │   └── main.py               # CLI con argparse (5 comandos)
│   ├── gui/
│   │   └── main.py               # GUI tkinter (4 pestañas)
│   └── utils/
│       └── logging.py            # Configuración de logging
├── tests/
│   ├── test_core/
│   └── test_integration/
├── config/
│   └── default.yaml              # Configuración por defecto
├── data/                         # Datos de entrada/salida
│   └── correos.xlsx
├── pyproject.toml
├── start.bat                     # Launcher Windows
├── setup_windows.bat             # Setup automático Windows
└── CLAUDE.md                     # Guía para IA
```

## Formato Excel

| Columna | Letra | Contenido |
|---------|-------|-----------|
| Correo | A | Email a procesar (input) |
| Status | B | OK / NO EXISTE / ERROR / vacío (=pendiente) |
| Nombre | C | Contact name |
| Email Personal | D | Email real del contacto |
| Teléfono | E | Número de trabajo |
| SIP | F | Dirección SIP |
| Dirección | G | Dirección postal |
| Departamento | H | Departamento/Unidad |
| Compañía | I | Organismo/Empresa |
| Oficina | J | Ubicación de oficina |

El sistema es **incremental**: solo procesa filas con Status vacío. Al terminar, escribe el resultado en la misma fila.

## Limitaciones Conocidas

- **Nombre bloqueado**: OWA anti-scraping impide extraer el nombre completo vía API y Playwright.
- **Sesión efímera**: ~40-50 llamadas API antes de que OWA exija re-autenticación (HTTP 307 → ADFS).
- **Anti-scraping server-side**: No se puede evadir con técnicas client-side (Patchright, stealth, etc.).

## Desarrollo

```bash
pip install -e ".[dev]"
pytest
ruff check src/
black src/ tests/
```
