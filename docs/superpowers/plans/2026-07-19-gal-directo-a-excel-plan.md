# GAL Directo a Excel - Plan de Implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Eliminar todos los JSON intermediarios. El Excel es la única fuente de verdad. GAL se extrae directo a Sheet1 del Excel, guardando tras cada página de API (~100 contactos). Enrichment lee de Sheet1 y actualiza la misma fila.

**Architecture:**
- `gal_scraper.py` → guarda directo a Sheet1 del Excel tras cada página de API (no guarda en JSON)
- `gal_exporter.py` → funciones de upsert por `persona_id` en Sheet1, mantiene Sheet2 para empresas
- `gal_enricher.py` → lee Sheet1 por `persona_id`, llama GetPersona, actualiza la misma fila
- JSONs eliminados: `gal_progress.json`, `gal_companies.json`, `directorio_completo.json`, `cache.json`, `enrich_progress.json`
- Se mantiene: `state.json` (sesión OWA)

**Tech Stack:** Python, openpyxl, OWA FindPeople API, GetPersona API

## Global Constraints

- Excel siempre guarda en disco tras cada página (no en memoria hasta el final)
- Resumable por `persona_id` en Sheet1 (unique key)
- Session expira → lo guardado ya está en Excel, se puede reanudar
- Mantener compatibilidad con el resto del código (config, GUI, etc.)

---

## Task 1: Refactor gal_exporter.py - Upsert por persona_id

**Files:**
- Modify: `src/verificacion_correo/core/gal_exporter.py`

**Interfaces:**
- Consumes: `flatten_contact_to_dict()` output format
- Produces: `append_contacts_to_excel(contacts, excel_path)` - upsert por persona_id

- [ ] **Step 1: Escribir test para upsert**

```python
# tests/test_core/test_gal_exporter.py
def test_append_contacts_upserts_by_persona_id():
    """Cuando el mismo persona_id existe, se actualiza; si no, se inserta."""
    from openpyxl import load_workbook
    from pathlib import Path
    import tempfile
    from verificacion_correo.core.gal_exporter import append_contacts_to_excel

    with tempfile.TemporaryDirectory() as tmpdir:
        excel_path = Path(tmpdir) / "test.xlsx"

        # Primera inserción
        contacts1 = [{
            'nombre': 'Juan', 'email': 'juan@test.com', 'empresa': 'Ayto',
            'telefono': '', 'departamento': '', 'oficina': '', 'direccion': '',
            'persona_id': 'abc123'
        }]
        append_contacts_to_excel(contacts1, excel_path)

        wb = load_workbook(excel_path)
        ws1 = wb["Contactos"]
        assert ws1.max_row == 2  # header + 1 data

        # Upsert con mismo persona_id pero datos diferentes
        contacts2 = [{
            'nombre': 'Juan', 'email': 'juan@test.com', 'empresa': 'Ayto',
            'telefono': '912345678', 'departamento': 'IT', 'oficina': 'Oficina 1',
            'direccion': 'Calle Mayor 1', 'persona_id': 'abc123'
        }]
        append_contacts_to_excel(contacts2, excel_path)

        wb = load_workbook(excel_path)
        ws1 = wb["Contactos"]
        assert ws1.max_row == 2  # sigue siendo 1 contacto (upsert)
        assert ws1.cell(2, 4).value == '912345678'  # teléfono actualizado
```

- [ ] **Step 2: Ejecutar test (debe fallar)**

Run: `cd /Users/andresgaibor/code/python/verificacion-correo && source .venv/bin/activate && python -m pytest tests/test_core/test_gal_exporter.py::test_append_contacts_upserts_by_persona_id -v`
Expected: FAIL (function not defined)

- [ ] **Step 3: Implementar append_contacts_to_excel**

Reemplazar el contenido de `gal_exporter.py` con:

```python
"""GAL Excel exporter con 2 hojas: Contactos + Compañías.

Sheet1 (Contactos) se actualiza con upsert por persona_id.
Sheet2 (Compañías) lista empresas extraídas para filtrar enrichment.
"""

from pathlib import Path
from typing import List, Dict, Any, Optional
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font

CONTACT_FIELDS = ['nombre', 'email', 'empresa', 'telefono', 'departamento', 'oficina', 'direccion', 'persona_id']


def flatten_contact_to_dict(persona: dict) -> dict:
    """Convierte persona del GAL a dict plano."""
    name = persona.get("DisplayName") or ""

    email_obj = persona.get("EmailAddress") or persona.get("EmailAddresses") or []
    email = ""
    if isinstance(email_obj, dict):
        email = email_obj.get("EmailAddress") or ""
    elif isinstance(email_obj, list) and email_obj:
        for e in email_obj:
            val = e.get("Value") or ''
            if val and '@' in val:
                email = val
                break
    elif isinstance(email_obj, str):
        email = email_obj

    phone = ""
    phone_items = persona.get("BusinessPhoneNumbersArray") or []
    if phone_items and isinstance(phone_items[0], dict):
        val = phone_items[0].get("Value") or {}
        if isinstance(val, dict):
            phone = val.get("Number") or val.get("NormalizedNumber") or ""
        elif isinstance(val, str):
            phone = val

    company = persona.get("CompanyName") or ""
    department = persona.get("Department") or ""
    office = persona.get("OfficeLocation") or ""

    address = ""
    addr_items = persona.get("BusinessAddressesArray") or []
    if addr_items and isinstance(addr_items[0], dict):
        val = addr_items[0].get("Value") or {}
        if isinstance(val, dict):
            parts = [val.get("Street") or "", val.get("City") or "", val.get("PostalCode") or "", val.get("State") or ""]
            parts = [p for p in parts if p.strip()]
            address = ", ".join(parts) if parts else ""
        elif isinstance(val, str):
            address = val

    persona_id = ""
    persona_id_obj = persona.get("PersonaId") or {}
    if isinstance(persona_id_obj, dict):
        persona_id = persona_id_obj.get("Id") or ""

    return {
        'nombre': name, 'email': email, 'empresa': company,
        'telefono': phone, 'departamento': department,
        'oficina': office, 'direccion': address, 'persona_id': persona_id,
    }


def _auto_width(ws):
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_length + 2, 40)


def append_contacts_to_excel(contacts: List[dict], excel_path: Path):
    """Upsert contacts en Sheet1 del Excel por persona_id.

    Sheet2 (Compañías) se recalcula automáticamente.
    Si el archivo no existe, lo crea con las 2 hojas.
    """
    excel_path = Path(excel_path)
    excel_path.parent.mkdir(parents=True, exist_ok=True)

    if excel_path.exists():
        wb = load_workbook(excel_path)
        ws1 = wb["Contactos"]
    else:
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Contactos"
        ws1.append(CONTACT_FIELDS)
        for cell in ws1[1]:
            cell.font = Font(bold=True)

    # Index existente por persona_id
    existing_rows: Dict[str, int] = {}
    for row_idx in range(2, ws1.max_row + 1):
        pid = ws1.cell(row_idx, 8).value  # persona_id es columna 8
        if pid:
            existing_rows[pid] = row_idx

    # Upsert contacts
    for contact in contacts:
        pid = contact.get('persona_id', '')
        if pid and pid in existing_rows:
            # Update existente
            row_idx = existing_rows[pid]
            for col_idx, field in enumerate(CONTACT_FIELDS, 1):
                val = contact.get(field, '')
                if val:  # solo actualizar si hay valor nuevo
                    ws1.cell(row_idx, col_idx).value = val
        else:
            # Insert nuevo
            row = [contact.get(f, '') for f in CONTACT_FIELDS]
            ws1.append(row)
            if pid:
                existing_rows[pid] = ws1.max_row

    # Recalcular Sheet2 con empresas únicas
    if "Compañías" in wb.sheetnames:
        del wb["Compañías"]
    ws2 = wb.create_sheet("Compañías")
    ws2.append(['Compañía', 'Enrich'])
    for cell in ws2[1]:
        cell.font = Font(bold=True)

    companies = set()
    for row_idx in range(2, ws1.max_row + 1):
        company = ws1.cell(row_idx, 3).value  # empresa es columna 3
        if company:
            companies.add(company)
    for company in sorted(companies):
        ws2.append([company, ''])

    _auto_width(ws1)
    _auto_width(ws2)
    wb.save(excel_path)
```

- [ ] **Step 4: Ejecutar test**

Run: `cd /Users/andresgaibor/code/python/verificacion-correo && source .venv/bin/activate && python -m pytest tests/test_core/test_gal_exporter.py::test_append_contacts_upserts_by_persona_id -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/andresgaibor/code/python/verificacion-correo
rtk git add src/verificacion_correo/core/gal_exporter.py tests/test_core/test_gal_exporter.py
rtk git commit -m "feat(gal_exporter): add append_contacts_to_excel with upsert by persona_id"
```

---

## Task 2: Refactor gal_scraper.py - directo a Excel

**Files:**
- Modify: `src/verificacion_correo/core/gal_scraper.py`

**Interfaces:**
- Consumes: `append_contacts_to_excel()` from gal_exporter
- Produces: `scrape_gal()` guarda directo a Excel tras cada página

- [ ] **Step 1: Modificar scrape_gal() para guardar a Excel tras cada página**

Eliminar:
- `PROGRESS_FILENAME = "gal_progress.json"`
- `OUTPUT_JSON = "directorio_completo.json"`
- `OUTPUT_CSV = "directorio_completo.csv"`
- `COMPANIES_CACHE_FILE = "gal_companies.json"`
- `save_to_json()` y `save_to_csv()`
- `ProgressFile` class

Mantener:
- `flatten_contact_to_dict()` y `append_contacts_to_excel()` para guardar a Excel

Cambios en `scrape_gal()`:
1. En vez de `all_people.extend(people)` y `progress.save()`, llamar `append_contacts_to_excel(people, excel_path)` tras cada página
2. Ya no guardar `scanned_all` (personas raw) - se leen del Excel si se reanuda
3. Para resume: leer Sheet1 del Excel, obtener max_row = offset inicial
4. Eliminar `output_excel` param (siempre se guarda a Excel)
5. Eliminar `json_path`, `csv_path` del resultado

```python
# Eliminar estas líneas de gal_scraper.py:
import csv  # ya no se necesita
PROGRESS_FILENAME = "gal_progress.json"
OUTPUT_JSON = "directorio_completo.json"
OUTPUT_CSV = "directorio_completo.csv"
COMPANIES_CACHE_FILE = "gal_companies.json"

# En scrape_gal(), donde antes hacía:
# all_people.extend(people)
# progress.save(offset, all_people)
# Ahora hacer:
from .gal_exporter import flatten_contact_to_dict, append_contacts_to_excel
flattened = [flatten_contact_to_dict(p) for p in people]
append_contacts_to_excel(flattened, excel_path)
```

**Resumen de cambios específicos en scrape_gal():**
- Línea ~390: eliminar `progress = ProgressFile(output_path)` y todas las llamadas a `progress.save()`
- Línea ~415: `all_people` ya no se acumula en memoria, solo se procesa cada página
- Línea ~520-523: donde dice `progress.save(scanned_offset, scanned_all)` → `append_contacts_to_excel(flattened_page, excel_path)`
- Línea ~670-675: eliminar `save_to_json(all_people, json_path)` y `save_to_csv(flattened, csv_path)`
- Función `save_to_json()` y `save_to_csv()` → eliminar
- Clase `ProgressFile` → eliminar
- Función `save_companies_cache()` y `load_companies_cache()` → mantener (se usan para el scan de compañías, no para scraping)

**Detalle del cambio en el loop principal (filtered mode, líneas ~500-530):**

Antes:
```python
scanned_all.extend(people)
scanned_offset += len(people)
progress.save(scanned_offset, scanned_all)
```

Después:
```python
flattened = [flatten_contact_to_dict(p) for p in people]
append_contacts_to_excel(flattened, excel_path)
scanned_offset += len(people)
```

**Para resume (recompatibilizado):** Leer `max_row` de Sheet1 del Excel para saber desde dónde continuar.

- [ ] **Step 2: Ejecutar test de sintaxis**

Run: `cd /Users/andresgaibor/code/python/verificacion-correo && source .venv/bin/activate && python -c "from verificacion_correo.core.gal_scraper import scrape_gal; print('OK')"`
Expected: OK (sin errores de sintaxis)

- [ ] **Step 3: Commit**

```bash
rtk git add src/verificacion_correo/core/gal_scraper.py
rtk git commit -m "refactor(gal_scraper): save to Excel after each page, remove JSON intermediates"
```

---

## Task 3: Refactor gal_enricher.py - leer de Excel, sin JSON

**Files:**
- Modify: `src/verificacion_correo/core/gal_enricher.py`

**Interfaces:**
- Consumes: `get_companies_to_enrich_from_excel()` (ya existe), Sheet1 del Excel
- Produces: Actualiza Sheet1 in-place

- [ ] **Step 1: Escribir test para enrichment desde Excel**

```python
# tests/test_core/test_gal_enricher.py
def test_enrich_reads_from_excel_not_json():
    """Enrichment debe leer contactos directamente del Excel, no de JSON cache."""
    # Verificar que enrich_excel_by_companies ya no acepta parametro cache
    import inspect
    from verificacion_correo.core.gal_enricher import enrich_excel_by_companies
    sig = inspect.signature(enrich_excel_by_companies)
    assert 'cache' not in sig.parameters, "enrich_excel_by_companies no debe aceptar parametro cache"
```

- [ ] **Step 2: Ejecutar test (debe fallar)**

Run: `cd /Users/andresgaibor/code/python/verificacion-correo && source .venv/bin/activate && python -m pytest tests/test_core/test_gal_enricher.py::test_enrich_reads_from_excel_not_json -v`
Expected: FAIL

- [ ] **Step 3: Modificar gal_enricher.py**

Eliminar:
- `EnrichProgress` class
- `find_contact_in_cache()` 
- `cache` param de `enrich_excel_by_companies()`
- `progress_path` param

Cambiar `enrich_excel_by_companies()`:
- Ya no recibe `cache: List[dict]` como param
- Busca `persona_id` directamente en Sheet1 (columna 8)
- Itera sobre Sheet1, para cada fila con `persona_id` y sin teléfono → llama GetPersona
- Solo procesa filas de companies marcadas con X en Sheet2

```python
# En gal_enricher.py - cambiar la función principal:

def enrich_excel_by_companies(
    excel_path: Path,
    companies: List[str],
    session_file: str = "state.json",
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> Dict[str, Any]:
    """Enriquece Excel selectivamente por lista de compañías usando GetPersona API.

    1. Lee Sheet2 -> filtra por companies con X
    2. Para cada compañía:
       a. Busca en Sheet1 filas donde Empresa == compañía
       b. Para cada match con persona_id -> llama GetPersona API
       c. Actualiza la misma fila en Sheet1
    3. Guarda Excel actualizado
    """
    companies_set = set(companies)

    try:
        cookie_str = _build_cookie_header(session_file)
        canary = _get_canary(session_file)
    except Exception:
        return {
            'companies_done': 0,
            'contacts_enriched': 0,
            'error': 'No se pudo cargar la sesión'
        }

    wb = load_workbook(excel_path)
    ws1 = wb["Contactos"]
    ws2 = wb["Compañías"]

    # Construir index de persona_id -> row para lookup rápido
    headers = [cell.value for cell in ws1[1]]
    persona_id_idx = headers.index('persona_id') + 1
    telefono_idx = headers.index('telefono') + 1
    depto_idx = headers.index('departamento') + 1
    oficina_idx = headers.index('oficina') + 1
    direccion_idx = headers.index('direccion') + 1
    empresa_idx = headers.index('empresa') + 1

    # Index: empresa -> lista de row_idx
    empresa_to_rows: Dict[str, List[int]] = {}
    for row_idx in range(2, ws1.max_row + 1):
        empresa = ws1.cell(row_idx, empresa_idx).value
        if empresa:
            empresa_to_rows.setdefault(empresa, []).append(row_idx)

    total_enriched = 0
    companies_done: List[str] = []
    errors: List[str] = []

    for company in companies:
        if company not in companies_set:
            continue

        row_indices = empresa_to_rows.get(company, [])
        for row_idx in row_indices:
            persona_id = ws1.cell(row_idx, persona_id_idx).value
            if not persona_id:
                continue

            # Solo enriquecer si teléfono está vacío
            telefono_actual = ws1.cell(row_idx, telefono_idx).value
            if telefono_actual:
                continue

            try:
                enriched = _call_get_persona(persona_id, cookie_str, canary)
            except Exception as e:
                errors.append(str(e))
                continue

            if not enriched:
                continue

            if enriched.get('telefono'):
                ws1.cell(row_idx, telefono_idx).value = enriched['telefono']
            if enriched.get('departamento'):
                ws1.cell(row_idx, depto_idx).value = enriched['departamento']
            if enriched.get('oficina'):
                ws1.cell(row_idx, oficina_idx).value = enriched['oficina']
            if enriched.get('direccion'):
                ws1.cell(row_idx, direccion_idx).value = enriched['direccion']

            total_enriched += 1
            if progress_callback:
                progress_callback(total_enriched, 0)

            time.sleep(0.5)

        companies_done.append(company)

    wb.save(excel_path)

    result = {
        'companies_done': len(companies_done),
        'contacts_enriched': total_enriched,
    }
    if errors:
        result['errors'] = errors[:5]

    return result
```

- [ ] **Step 4: Ejecutar test**

Run: `cd /Users/andresgaibor/code/python/verificacion-correo && source .venv/bin/activate && python -m pytest tests/test_core/test_gal_enricher.py::test_enrich_reads_from_excel_not_json -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
rtk git add src/verificacion_correo/core/gal_enricher.py tests/test_core/test_gal_enricher.py
rtk git commit -m "refactor(gal_enricher): read from Excel Sheet1, remove JSON cache dependency"
```

---

## Task 4: Actualizar gui/service.py

**Files:**
- Modify: `src/verificacion_correo/gui/service.py`

**Interfaces:**
- Consumes: `append_contacts_to_excel()` output
- Produces: `start_gal_scraping()` y `start_enrichment()` actualizados

- [ ] **Step 1: Actualizar start_enrichment()**

Cambiar línea ~190-245 para:
1. Ya no llamar `load_raw_gal_cache()` (eliminar)
2. Ya no pasar `cache` a `enrich_excel_by_companies()`
3. Ya no pasar `progress_path`

```python
# En service.py - start_enrichment()
def start_enrichment(self, excel_path: str) -> None:
    # ...
    def enrich_thread():
        try:
            from verificacion_correo.core.gal_enricher import (
                enrich_excel_by_companies,
                get_companies_to_enrich_from_excel,
            )

            excel_p = Path(excel_path)
            companies = get_companies_to_enrich_from_excel(excel_p)

            if not companies:
                self.progress_queue.put(('enrich_complete', {
                    'error': 'No companies selected for enrichment',
                    'contacts_enriched': 0,
                    'companies_done': 0
                }))
                self.is_processing = False
                return

            def progress_callback(enriched_count, total):
                self.progress_queue.put(('enrich_progress', {
                    'count': enriched_count,
                    'companies': len(companies)
                }))

            result = enrich_excel_by_companies(
                excel_p,
                companies,
                progress_callback=progress_callback,
            )
            self.progress_queue.put(('enrich_complete', result))
```

- [ ] **Step 2: Ejecutar test de sintaxis**

Run: `cd /Users/andresgaibor/code/python/verificacion-correo && source .venv/bin/activate && python -c "from verificacion_correo.gui.service import GUIService; print('OK')"`
Expected: OK

- [ ] **Step 3: Commit**

```bash
rtk git add src/verificacion_correo/gui/service.py
rtk git commit -m "refactor(gui_service): start_enrichment reads from Excel only"
```

---

## Task 5: Actualizar gui/main.py

**Files:**
- Modify: `src/verificacion_correo/gui/main.py`

- [ ] **Step 1: Cambiar nombre de archivo gal_export.xlsx → gal_directorio.xlsx**

Buscar y reemplazar en main.py:
- `gal_export.xlsx` → `gal_directorio.xlsx`

Líneas afectadas (~1074-1075):
```python
excel_path = output_dir / "gal_directorio.xlsx"
```

También en `_start_scraper()` alrededor de línea 1120 donde llama `start_gal_scraping`.

- [ ] **Step 2: Eliminar checks de gal_progress.json**

En `_start_scraper()` (líneas ~1076-1090), eliminar el bloque que pregunta si quiere resumir basándose en `gal_progress.json`. Ahora el resume se hace leyendo el Excel directamente.

Reemplazar con:
```python
# Verificar si existe Excel previo
if excel_path.exists():
    resume = messagebox.askyesno(
        "Reanudar Extracción",
        f"Se encontró una extracción anterior:\n{excel_path}\n\n"
        f"¿Desea continuar desde donde se quedó?\n\n"
        f"Seleccione 'No' para empezar desde cero (sobrescribirá)."
    )
    force_restart = not resume
else:
    force_restart = True
```

- [ ] **Step 3: Commit**

```bash
rtk git add src/verificacion_correo/gui/main.py
rtk git commit -m "refactor(gui): use gal_directorio.xlsx, resume from Excel"
```

---

## Task 6: Limpiar archivos JSON obsoletos

**Files:**
- Delete: `data/gal/gal_progress.json` (si existe)
- Delete: `data/gal/enrich_progress.json` (si existe)
- Modify: `.gitignore` (si no está ya)

- [ ] **Step 1: Asegurar que data/ está en .gitignore**

Verificar que `.gitignore` contiene `data/` (ya debe estar según work anterior).

- [ ] **Step 2: Commit**

```bash
rtk git add -A
rtk git commit -m "chore: remove stale JSON intermediates from git"
```

---

## Verificación Final

- [ ] `python -c "from verificacion_correo.core.gal_scraper import scrape_gal; from verificacion_correo.core.gal_exporter import append_contacts_to_excel; from verificacion_correo.core.gal_enricher import enrich_excel_by_companies; print('All imports OK')"`
- [ ] `python -m pytest tests/test_core/test_gal_exporter.py tests/test_core/test_gal_enricher.py -v`
- [ ] Verificar que no hay archivos `.json` de cache en el código (solo `state.json` para sesión)
