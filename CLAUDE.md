# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python automation tool that uses Playwright to interact with the Madrid city webmail interface (correoweb.madrid.org/owa). The application automates the process of extracting contact information from email recipients by interacting with the OWA (Outlook Web Access) interface.

## Architecture

El proyecto está organizado en módulos con responsabilidades separadas:

### Módulos Principales

1. **copiar_sesion.py** - Gestión de Sesión
   - Lanza navegador para autenticación manual
   - Guarda el estado de autenticación en `state.json` para reutilización
   - Debe ejecutarse primero para establecer una sesión válida

2. **config.py** - Configuración Centralizada
   - `PAGE_URL`: URL de OWA
   - `DEFAULT_EMAILS`: Emails de respaldo si no existe archivo Excel
   - Patrones Regex: EMAIL_RE, PHONE_RE, POSTAL_ADDR_RE, SIP_RE, NAME_RE
   - `SELECTORS`: Diccionario con selectores CSS para elementos de la interfaz OWA
   - `WAIT_TIMES`: Tiempos de espera configurables (milisegundos)
   - `BROWSER_CONFIG`: Configuración del navegador (headless, session_file)
   - `EXCEL_CONFIG`: Configuración de lectura de Excel (archivo, fila inicial, columna)

3. **excel_reader.py** - Lectura de Datos
   - `leer_correos_excel(archivo_path)`: Lee emails desde archivo Excel
     - Lee columna A desde fila 2 en adelante (configurable)
     - Retorna lista de direcciones de email
     - Fallback a DEFAULT_EMAILS si archivo no existe o está vacío

4. **contact_extractor.py** - Extracción de Información
   - `extract_from_popup_text(text)`: Extracción basada en regex del texto del popup
   - `popup_info(page)`: Función principal de extracción que:
     1. Espera popup con selector `div._pe_Y[ispopup="1"]`
     2. Intenta selectores DOM específicos (ej: `span._pe_c1._pe_t1` para nombre)
     3. Fallback a extracción regex desde texto completo
     4. Consolida resultados priorizando selectores específicos sobre regex matches

5. **browser_automation.py** - Automatización del Navegador
   - `procesar_emails(emails)`: Función principal que orquesta el proceso completo
     - Lanza navegador con sesión guardada
     - Navega a OWA y abre nuevo mensaje
     - Procesa cada email: hace clic en token, extrae info, cierra popup
     - Retorna lista de resultados

6. **app.py** - Script Principal (Punto de Entrada)
   - Script simplificado (~30 líneas) que orquesta todo el proceso
   - Lee emails desde Excel usando `excel_reader`
   - Procesa emails usando `browser_automation`
   - Muestra resumen final

## Development Setup

1. **Create virtual environment**:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

2. **Install dependencies**:
   ```bash
   pip install playwright openpyxl
   playwright install chromium
   ```

3. **Configure email addresses**:
   - Create or edit `data/correos.xlsx`
   - Place email addresses in column A, starting from row 2
   - Row 1 should contain the header "Correo"
   - Example structure:
     ```
     | A                 |
     |-------------------|
     | Correo            | <- Row 1 (header)
     | user1@madrid.org  | <- Row 2
     | user2@madrid.org  | <- Row 3
     | ...               |
     ```
   - An example file is provided with default emails

4. **Create session file**:
   ```bash
   python copiar_sesion.py
   # Follow the manual login prompts, then press ENTER to save session
   ```

5. **Run the automation**:
   ```bash
   python app.py
   ```

## Important Implementation Details

- **Session Persistence**: The `state.json` file contains authentication cookies and must exist before running `app.py`
- **DOM Selectors**: The script targets specific OWA interface elements:
  - New message button: `button[title="Escribir un mensaje nuevo (N)"]`
  - To field: `textbox` with role and name "Para"
  - Email tokens: Located via regex pattern matching on visible text
  - Popup cards: `div._pe_Y[ispopup="1"]`
- **Timing Strategy**: Uses explicit waits (`wait_for_timeout`, `wait_for_selector`) to handle dynamic content loading
- **Error Handling**: Uses try-except blocks for optional extractions (see `safe_text()` and popup selector waits)
- **Contact Extraction Strategy**:
  - Primary: Use specific DOM selectors (class names, autoids)
  - Fallback: Extract from raw text using regex
  - Name format expected: "APELLIDO, NOMBRE" (surname, firstname)
  - Office/job title: Heuristic looks for all-uppercase lines

## Common Patterns

- **Launching browser with saved session**:
  ```python
  browser = p.chromium.launch(headless=False)
  context = browser.new_context(storage_state="state.json")
  ```

- **Safe element extraction**:
  ```python
  def safe_text(locator):
      try:
          return locator.inner_text(timeout=1000).strip()
      except:
          return None
  ```

- **Filtering locators by email list**:
  ```python
  escaped_emails = [re.escape(email) for email in emails]
  pattern = re.compile(r"^(?:" + "|".join(escaped_emails) + r")$", re.IGNORECASE)
  email_tokens = email_tokens.filter(has_text=pattern)
  ```

## Key Dependencies

- **playwright**: Browser automation framework (requires `playwright install` after pip install)
- **openpyxl**: Library for reading/writing Excel files (.xlsx format)
- Python 3.13 (as indicated by venv structure)

## Project Structure

```
verificacion-correo/
├── config.py                 # Configuración centralizada
├── excel_reader.py           # Lectura de emails desde Excel
├── contact_extractor.py      # Extracción de información de contacto
├── browser_automation.py     # Automatización del navegador
├── app.py                    # Script principal (punto de entrada)
├── copiar_sesion.py          # Gestión de sesión
├── examples/                 # Scripts de referencia y debugging
│   ├── app_auto.py          # Versión monolítica con output detallado
│   └── app_debug.py         # Script de debugging
├── data/                     # Archivos de datos
│   └── correos.xlsx         # Excel con emails a procesar
├── state.json                # Sesión guardada (ignorar en git)
└── .venv/                    # Entorno virtual (ignorar en git)
```

## Files to Ignore

- `state.json`: Contains session authentication data (already in .gitignore)
- `.venv/`: Virtual environment directory
- `data/`: Directory for storing data files including:
  - `correos.xlsx`: Excel file containing email addresses to process (example file provided)
- `examples/`: Scripts de referencia (app_auto.py, app_debug.py) y pruebas antiguas
- `.chrome_user_data/`: Directorio de usuario de Chrome para Patchright (si existe)

## Limitaciones Conocidas - Microsoft OWA Anti-Scraping

### Protección del Nombre de Usuario

Microsoft OWA (Outlook Web Access) implementa medidas anti-scraping robustas que **previenen específicamente la extracción del nombre completo** del contacto cuando se detecta automatización.

**Síntomas**:
- El popup de tarjeta de contacto muestra el email del token (ej: "ASP164@MADRID.ORG") en lugar del nombre real ("SERRANO PEREZ, ANTONIO MANUEL")
- Todos los demás campos se cargan correctamente

**Técnicas probadas SIN ÉXITO**:
1. ✗ Playwright básico con configuración stealth
2. ✗ **Patchright** - La librería anti-detección más avanzada (2025)
   - Parchea Playwright a nivel de código fuente
   - Evita fugas CDP (Chrome DevTools Protocol)
   - channel="chrome", headless=False, no_viewport=True
   - **Resultado**: OWA sigue detectando la automatización

**Conclusión**: Microsoft OWA tiene protección anti-bot a nivel del servidor que no se puede evadir con técnicas del lado del cliente.

### Datos que SÍ se Extraen Correctamente

El script extrae exitosamente **8 de 9 campos**:
- ✅ Email personal (ej: `antoniomanuel.serrano@madrid.org`)
- ✅ Teléfono de trabajo (ej: `916704092`)
- ✅ Dirección completa (ej: `C/ AYUNTAMIENTO, 5 28791 RIVAS-VACIAMADRID MADRID`)
- ✅ Departamento (ej: `OFICINA JUDICIAL MUNICIPAL`)
- ✅ Compañía (ej: `ORGANOS JUDICIALES`)
- ✅ Ubicación de oficina (ej: `RIVAS-VACIAMADRID`)
- ✅ Dirección SIP (ej: `sip:asp164@madrid.org`)
- ✅ Token email (para identificación)
- ❌ **Nombre completo** (bloqueado por OWA anti-scraping)

### Alternativas para Obtener Nombres

Si necesitas los nombres completos:
1. Enriquecer los datos con una fuente externa (Active Directory, lista de empleados, etc.)
2. Obtener los nombres manualmente de otra interfaz de OWA
3. Usar el email personal para inferir el nombre (ej: `antoniomanuel.serrano@` → Antonio Manuel Serrano)
