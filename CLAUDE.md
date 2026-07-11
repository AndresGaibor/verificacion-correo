# CLAUDE.md — Guía para IA

## Proyecto

Automatización Python que extrae contactos del directorio OWA de Madrid (correoweb.madrid.org) usando la API REST de Exchange (FindPeople + GetPersona). También incluye un scraper vía Playwright como alternativa legacy.

## Arquitectura (`src/`)

### Core

- **api_extractor.py** — Extracción vía API REST (recomendado).
  - `_find_persona_id()`: busca en GAL por email → obtiene PersonaId
  - `_get_persona()`: obtiene datos completos (phone, address, SIP, dept, etc.)
  - `_parse_persona()`: convierte respuesta JSON a `ContactInfo`
  - `validate_session_api()`: validación ligera sin Playwright (FindPeople sin query)
  - `SessionExpiredError`: se lanza cuando OWA responde HTTP 307 (sesión expirada)
  - `process_emails_via_api()`: orquesta procesamiento batch completo
  - `SESSION_HEALTH_CHECK_INTERVAL = 5`, `SESSION_ESTIMATED_LIMIT = 40` — checks each 5 calls
  - Usa `urllib` + cookies de sesión (`state.json`), sin Playwright
  - Delay entre requests: 3s
  - Timeout: 60s

- **gal_scraper.py** — Extracción completa del GAL (Global Address List).
  - `scrape_gal()`: paginación con Offset, reanudable vía `ProgressFile`
  - `company_filter` param: filtrado server-side via AQS QueryString en FindPeople
  - Cuando `company_filter` está activo, itera compañía por compañía (más rápido, menos carga de red)
  - `enrich_contacts` param: llama a GetPersona por cada contacto para datos completos (phone, dept, address)
  - `ProgressFile`: guarda offset + contador + `completed_companies` para resumir por compañía
  - `_flatten_persona()`: aplanado para CSV
  - `_enrich_persona()`: enriquece contacto vía GetPersona (phone, department, office, address)
  - `_build_find_people_payload()`: construye payload JSON con `query_string` para QueryString AQS
  - Exporta a JSON + CSV (delimitador `;`)
  - Delay configurable (default 8s), batch size 100
  - Acepta `stop_flag` (dict con `{"stop": bool}`) para parada externa
  - Acepta `progress_callback(count, total)` para UI
  - SessionExpiredError: guarda progreso y termina ordenadamente

- **browser.py** — Automatización con Playwright (legacy).
  - `BrowserAutomation`: clase principal, orquesta el proceso
  - `process_emails()`: lee Excel, procesa batches, escribe resultados
  - Threa-safe: remueve event loop de asyncio antes de Playwright
  - `_process_batch()`: abre mensaje nuevo, añade destinatarios, procesa cada email
  - `_process_single_email()`: hace clic en token, espera popup, extrae info
  - `_find_email_token()`: busca token por `img[src]` o `title`/`inner_text`
  - `_is_valid_contact()`: validación multi-criterio del contacto extraído

- **extractor.py** — Extracción de contactos desde popups OWA.
  - `ContactInfo`: dataclass con name, email, phone, sip, address, department, company, office_location
  - `ContactExtractor`: multi-capa (DOM → texto → regex)
  - `_extract_by_text_labels()`: busca etiquetas fijas (Departamento:, Compañía:, Trabajo:, MI:)
  - `_extract_personal_email()`: filtra tokens genéricos (ASP123@, AGM456@)
  - Guarda screenshots y HTML de popups en `debug_screenshots/`

- **config.py** — Config centralizada via YAML.
  - `Config`: carga de `config/default.yaml` con fallback a `config.yaml`
  - Dataclasses: `BrowserConfig`, `ExcelConfig`, `ProcessingConfig`, `Selectors`, `WaitTimes`, `RegexPatterns`
  - `get_config()`: singleton global
  - Soporta PyInstaller (archivos embebidos en `sys._MEIPASS`)

- **excel.py** — Operaciones Excel incrementales.
  - `ExcelReader.read_pending_emails()`: lee filas con Status vacío → `ExcelSummary` con batches
  - `ExcelWriter.write_result()` / `write_batch_results()`: escribe Status + datos de contacto
  - `ProcessingStatus`: PENDING(""), SUCCESS("OK"), NOT_FOUND("NO EXISTE"), ERROR("ERROR")
  - Columnas fijas A-J (ver README.md)
  - Auto-creación del Excel con datos de ejemplo si no existe

- **session.py** — Gestión de sesión del navegador.
  - `SessionManager`: creación, validación, limpieza
  - `setup_interactive_session()`: abre navegador visible para login manual (5 min timeout)
  - `create_automation_context()`: contexto con `storage_state` cargado
  - `validate_session()`: navega a OWA, busca botón de nuevo mensaje
  - Threa-safe (remueve event loop como browser.py)

- **first_run.py** — Setup inicial automático.
  - `FirstRunManager`: crea config, directorios, Excel y Playwright browsers
  - `install_playwright_browsers()`: subprocess a `playwright install chromium`
  - `check_and_run_first_time_setup()`: entry point unificado

### CLI (`cli/main.py`)

- `VerificacionCorreoCLI`: 5 comandos con argparse
  - `process`: procesa pendientes (Playwright)
  - `setup`: configura sesión
  - `validate`: valida setup
  - `status`: muestra estado
  - `scrape-gallery`: GAL vía API, con --output-dir, --max-contacts, --batch-size, --delay, --force-restart, --company-filter, --enrich
    - `--company-filter`: descarga todo el GAL y filtra client-side por CompanyName (Exchange OWA FindPeople NO soporta filtrado server-side por compañía — Restriction devuelve HTTP 500)
    - `--enrich`: tras filtrar, enriquecen los resultados con GetPersona (teléfono, departamento, dirección)
- `__main__.py` → `python -m verificacion_correo` lanza CLI

### GUI (`gui/main.py`)

- Tkinter, 4 pestañas (Procesamiento, Sesión, Configuración, Scraper de Contactos)
- `GUIService`: orquesta procesamiento/API/GAL en background thread
- Scraper tab: usa `api_extractor.py` + `gal_scraper.py` directamente (sin Playwright)

## Entry Points (pyproject.toml)

```
verificacion-correo = "verificacion_correo.cli.main:main"
vcorreo = "verificacion_correo.cli.main:main"
verificacion-correo-gui = "verificacion_correo.gui.main:main"
```

## Formato Excel

| Col | Header | Descripción |
|-----|--------|-------------|
| A | Correo | Email a procesar |
| B | Status | OK / NO EXISTE / ERROR / "" |
| C-J | Datos | Nombre, Email Personal, Teléfono, SIP, Dirección, Departamento, Compañía, Oficina |

Sistema incremental: solo procesa filas con Status vacío. `ExcelColumns` clase con definiciones.

## Sesión OWA

- Guardar: `verificacion-correo setup` → login manual en navegador
- Archivo: `state.json` (Playwright storage_state con cookies)
- La sesión expira tras ~40-50 llamadas API → HTTP 307 → `SessionExpiredError`
- Health check cada 5 llamadas (`SESSION_HEALTH_CHECK_INTERVAL = 5`)
- `_build_cookie_header()` extrae cookies del state.json
- `_get_canary()` extrae `X-OWA-CANARY` de cookies o localStorage

## Patrones Importantes

### Procesamiento API (api_extractor.py)

```python
from verificacion_correo.core.api_extractor import find_people, process_emails_via_api

# Por email
contact = find_people("user@madrid.org", "state.json")

# Batch desde Excel
stats = process_emails_via_api("data/correos.xlsx", "state.json", batch_size=10)
```

### GAL Scraper

```python
from verificacion_correo.core.gal_scraper import scrape_gal

stats = scrape_gal("state.json", output_dir="data/gal", max_contacts=0, batch_size=100, company_filter=["ORGANOS JUDICIALES"], enrich_contacts=True)
```

### Sesión expirada

```python
from verificacion_correo.core.api_extractor import SessionExpiredError

try:
    contact = find_people(email, session_file)
except SessionExpiredError:
    # guardar progreso, pedir re-autenticación
    break
```

### Thread safety para Playwright

```python
import asyncio
old_loop = asyncio.get_event_loop()
asyncio.set_event_loop(None)
# ... operaciones Playwright ...
asyncio.set_event_loop(old_loop)
```

## Limitaciones Conocidas

- **Nombre bloqueado**: OWA anti-scraping server-side impide extraer DisplayName. Intentos con Patchright/stealth fallaron.
- **Sesión caduca**: ~40 llamadas API antes de HTTP 307. El GAL scraper es reanudable.
- **AddressListId fijo**: `fed75805-8ba2-4323-9f6d-80be7e3abc6a` para Madrid. Puede variar por organización.

## Dependencias

- playwright (solo para browser legacy)
- openpyxl (lectura/escritura Excel)
- pyyaml (config)
- Python >= 3.8

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
playwright install chromium
```

## Tests

```bash
pytest
pytest --cov=src --cov-report=html
```

## Archivos Legacy (no tocar)

Los siguientes archivos en la raíz son legacy v1 y ya no se usan. El código activo está en `src/`:

- `app.py`, `browser_automation.py`, `contact_extractor.py`, `config.py`, `excel_reader.py`, `excel_writer.py`, `copiar_sesion_old.py`, `debug_scraper.py`, `windows_compat.py`, `build_test.py`, `examples_old/`
