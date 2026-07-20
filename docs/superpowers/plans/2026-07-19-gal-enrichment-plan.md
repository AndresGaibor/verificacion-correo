# GAL Extracción + Enriquecimiento por Compañía

**Goal:** Sistema de extracción GAL que exporta a Excel (2 hojas) y permite enrichment selectivo por compañía.

**Architecture:** Extracción descarga GAL completo a JSON cache + Excel. Enrichment usa cache local para llenar campos vacíos solo de contactos que hagan match exacto por compañía.

**Tech Stack:** Python, openpyxl, gal_scraper.py existente

---

## Files

| Archivo | Responsabilidad |
|---------|-----------------|
| `src/verificacion_correo/core/gal_exporter.py` | Excel 2 hojas + cache JSON |
| `src/verificacion_correo/core/gal_enricher.py` | Logic de enrichment con progress |
| `src/verificacion_correo/gui/main.py` | UI con botón enrich + Sheet 2 |

---

## Task 1: Crear gal_exporter.py con Excel 2 hojas

**Files:**
- Create: `src/verificacion_correo/core/gal_exporter.py`
- Modify: `src/verificacion_correo/core/__init__.py` (exportar nuevas funciones)
- Test: `tests/test_core/test_gal_exporter.py`

**Interfaces:**
- Consumes: `List[dict]` de contactos flatten (del GAL)
- Produces: Excel en `output_path` con 2 hojas

- [ ] **Step 1: Crear estructura básica de gal_exporter.py**

```python
"""GAL Excel exporter con 2 hojas: Contactos + Compañías."""

from pathlib import Path
from typing import List, Dict, Any, Optional
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

CONTACT_FIELDS = ['nombre', 'email', 'empresa', 'telefono', 'departamento', 'oficina', 'direccion']

def flatten_contact_to_dict(persona: dict) -> dict:
    """Convierte persona del GAL a dict plano."""
    return {
        'nombre': persona.get('DisplayName', ''),
        'email': _get_email(persona),
        'empresa': persona.get('CompanyName', ''),
        'telefono': '',
        'departamento': '',
        'oficina': '',
        'direccion': '',
    }

def _get_email(persona: dict) -> str:
    """Extrae email principal del persona."""
    emails = persona.get('EmailAddresses') or []
    for e in emails:
        val = e.get('Value') or ''
        if val and '@' in val:
            return val
    return ''

def extract_companies_from_contacts(contacts: List[dict]) -> List[str]:
    """Extrae lista única de compañías de contactos."""
    companies = set()
    for c in contacts:
        company = c.get('empresa', '').strip()
        if company:
            companies.add(company)
    return sorted(companies)

def save_to_excel(contacts: List[dict], output_path: Path, cache_path: Optional[Path] = None):
    """Guarda contactos en Excel de 2 hojas.

    Sheet1: Contactos con todos los campos
    Sheet2: Compañías con checkbox Enrich
    """
    wb = Workbook()

    # Sheet 1: Contactos
    ws1 = wb.active
    ws1.title = "Contactos"
    headers = CONTACT_FIELDS
    ws1.append(headers)
    for contact in contacts:
        row = [contact.get(f, '') for f in fields]
        ws1.append(row)

    # Sheet 2: Compañías
    ws2 = wb.create_sheet("Compañías")
    ws2.append(['Compañía', 'Enrich'])
    companies = extract_companies_from_contacts(contacts)
    for company in companies:
        ws2.append([company, ''])

    wb.save(output_path)

    # Guardar cache JSON si se especifica
    if cache_path:
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(contacts, f, ensure_ascii=False, indent=2)

def load_gal_cache(cache_path: Path) -> List[dict]:
    """Carga GAL cache desde JSON."""
    with open(cache_path, 'r', encoding='utf-8') as f:
        return json.load(f)
```

- [ ] **Step 2: Escribir tests para gal_exporter**

```python
import pytest
from pathlib import Path
from verificacion_correo.core.gal_exporter import (
    flatten_contact_to_dict,
    extract_companies_from_contacts,
    save_to_excel,
    load_gal_cache,
)

def test_flatten_contact():
    persona = {
        'DisplayName': 'Juan Perez',
        'EmailAddresses': [{'Value': 'juan@madrid.org'}],
        'CompanyName': 'ORGANOS JUDICIALES',
    }
    result = flatten_contact_to_dict(persona)
    assert result['nombre'] == 'Juan Perez'
    assert result['email'] == 'juan@madrid.org'
    assert result['empresa'] == 'ORGANOS JUDICIALES'

def test_extract_companies():
    contacts = [
        {'empresa': 'ORG1'},
        {'empresa': 'ORG2'},
        {'empresa': 'ORG1'},
    ]
    result = extract_companies_from_contacts(contacts)
    assert result == ['ORG1', 'ORG2']

def test_save_to_excel_creates_two_sheets(tmp_path):
    contacts = [
        {'nombre': 'Test', 'email': 'test@test.com', 'empresa': 'TEST CO'},
    ]
    excel_path = tmp_path / "test.xlsx"
    save_to_excel(contacts, excel_path)

    from openpyxl import load_workbook
    wb = load_workbook(excel_path)
    assert "Contactos" in wb.sheetnames
    assert "Compañías" in wb.sheetnames

    ws2 = wb["Compañías"]
    assert ws2.cell(2, 1).value == "TEST CO"
```

- [ ] **Step 3: Correr tests**

Run: `pytest tests/test_core/test_gal_exporter.py -v`
Expected: PASS

- [ ] **Step 4: Commit**

---

## Task 2: Crear gal_enricher.py con enrichment selectivo

**Files:**
- Create: `src/verificacion_correo/core/gal_enricher.py`
- Test: `tests/test_core/test_gal_enricher.py`

**Interfaces:**
- Consumes: `excel_path`, `companies_to_enrich`, `cache`
- Produces: Excel actualizado con campos llenos solo donde vacíos

```python
"""GAL enrichment selectivo por compañía desde cache local."""

from pathlib import Path
from typing import List, Dict, Any, Set, Optional, Callable
import json
import time
from openpyxl import load_workbook

CONTACT_FIELDS = ['nombre', 'email', 'empresa', 'telefono', 'departamento', 'oficina', 'direccion']

class EnrichProgress:
    """Trackea progreso de enrichment para reanudación."""

    def __init__(self, path: Path):
        self.path = path
        self.data = {
            'companies_done': [],
            'contacts_enriched': 0,
            'offset': 0,
        }

    def save(self):
        with open(self.path, 'w') as f:
            json.dump(self.data, f, indent=2)

    def load(self) -> bool:
        if self.path.exists():
            with open(self.path, 'r') as f:
                self.data = json.load(f)
            return True
        return False

def find_contact_in_cache(email: str, empresa: str, cache: List[dict]) -> Optional[dict]:
    """Busca contacto exacto en cache por email + empresa."""
    for c in cache:
        if c.get('email', '').lower() == email.lower() and c.get('empresa', '') == empresa:
            return c
    return None

def merge_enrichment(existing: dict, enrichment: dict) -> dict:
    """Mezcla enrichment en existing - solo llena vacíos."""
    result = existing.copy()
    for key in ['telefono', 'departamento', 'oficina', 'direccion']:
        if not result.get(key) and enrichment.get(key):
            result[key] = enrichment.get(key)
    return result

def enrich_excel_by_companies(
    excel_path: Path,
    companies: List[str],
    cache: List[dict],
    progress_path: Optional[Path] = None,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> Dict[str, Any]:
    """Enriquece Excel selectivamente por lista de compañías.

    1. Lee Sheet2 -> filtra por companies con X
    2. Para cada compañía:
       a. Busca en Sheet1 filas donde Empresa == compañía
       b. Para cada match -> lookup cache -> llena vacíos
       c. Cada 50 contactos -> guarda progress
    3. Guarda Excel actualizado
    """
    companies_set = set(companies)
    progress = EnrichProgress(progress_path) if progress_path else None
    if progress and progress.load():
        companies_set -= set(progress.data['companies_done'])

    wb = load_workbook(excel_path)
    ws1 = wb["Contactos"]
    ws2 = wb["Compañías"]

    # Construir índice de cache por (email_lower, empresa)
    cache_index = {}
    for c in cache:
        key = (c.get('email', '').lower(), c.get('empresa', ''))
        cache_index[key] = c

    # Headers de Sheet1
    headers = [cell.value for cell in ws1[1]]
    telefono_idx = headers.index('telefono') + 1
    depto_idx = headers.index('departamento') + 1
    oficina_idx = headers.index('oficina') + 1
    direccion_idx = headers.index('direccion') + 1

    total_enriched = 0
    companies_done = []

    # Iterar Sheet2 para encontrar empresas marcadas
    for row_idx in range(2, ws2.max_row + 1):
        company = ws2.cell(row_idx, 1).value
        enrich_mark = ws2.cell(row_idx, 2).value

        if not company or enrich_mark != 'X':
            continue
        if company not in companies_set:
            continue

        # Buscar contactos en Sheet1 que matcheen esta compañía
        for data_row_idx in range(2, ws1.max_row + 1):
            empresa_cell = ws1.cell(data_row_idx, headers.index('empresa') + 1).value
            if empresa_cell != company:
                continue

            email_cell = ws1.cell(data_row_idx, headers.index('email') + 1).value
            cache_key = (email_cell.lower() if email_cell else '', company)
            cached = cache_index.get(cache_key)

            if not cached:
                continue

            # Solo llenar vacíos
            telefono = ws1.cell(data_row_idx, telefono_idx).value
            if not telefono and cached.get('telefono'):
                ws1.cell(data_row_idx, telefono_idx).value = cached['telefono']

            depto = ws1.cell(data_row_idx, depto_idx).value
            if not depto and cached.get('departamento'):
                ws1.cell(data_row_idx, depto_idx).value = cached['departamento']

            oficina = ws1.cell(data_row_idx, oficina_idx).value
            if not oficina and cached.get('oficina'):
                ws1.cell(data_row_idx, oficina_idx).value = cached['oficina']

            direccion = ws1.cell(data_row_idx, direccion_idx).value
            if not direccion and cached.get('direccion'):
                ws1.cell(data_row_idx, direccion_idx).value = cached['direccion']

            total_enriched += 1

            if progress_callback:
                progress_callback(total_enriched, 0)

        companies_done.append(company)

        # Guardar progress cada vez que se completa una compañía
        if progress:
            progress.data['companies_done'] = companies_done
            progress.data['contacts_enriched'] = total_enriched
            progress.save()

    # Guardar Excel
    wb.save(excel_path)

    return {
        'companies_done': len(companies_done),
        'contacts_enriched': total_enriched,
    }
```

- [ ] **Step 1: Crear gal_enricher.py con el código arriba**

- [ ] **Step 2: Escribir tests básicos**

```python
def test_merge_enrichment_fills_empty():
    existing = {'telefono': '', 'departamento': 'IT'}
    enrichment = {'telefono': '123456', 'departamento': ''}
    result = merge_enrichment(existing, enrichment)
    assert result['telefono'] == '123456'
    assert result['departamento'] == 'IT'

def test_merge_enrichment_does_not_overwrite():
    existing = {'telefono': '111', 'departamento': 'IT'}
    enrichment = {'telefono': '222', 'departamento': 'Sales'}
    result = merge_enrichment(existing, enrichment)
    assert result['telefono'] == '111'
    assert result['departamento'] == 'IT'
```

- [ ] **Step 3: Correr tests**

Run: `pytest tests/test_core/test_gal_enricher.py -v`

- [ ] **Step 4: Commit**

---

## Task 3: Modificar GUI - cambiar checkbox a botón y restaurar Sheet 2

**Files:**
- Modify: `src/verificacion_correo/gui/main.py:546-551` (quitar checkbox enrich)
- Modify: `src/verificacion_correo/gui/main.py:554-580` (agregar botón enrich)
- Modify: `src/verificacion_correo/gui/service.py` (agregar método enrich)

**Cambios en service.py:**
```python
def start_enrichment(self, excel_path: str, cache_path: str) -> None:
    """Inicia enrichment en thread background."""
    def enrich_thread():
        from verificacion_correo.core.gal_enricher import enrich_excel_by_companies
        from verificacion_correo.core.gal_exporter import load_gal_cache

        cache = load_gal_cache(Path(cache_path))

        # Leer empresas marcadas del Excel
        from openpyxl import load_workbook
        wb = load_workbook(excel_path)
        ws2 = wb["Compañías"]
        companies = []
        for row in ws2.iter_rows(min_row=2, values_only=True):
            if row[1] == 'X' and row[0]:
                companies.append(row[0])

        if not companies:
            self.progress_queue.put(('enrich_complete', {'error': 'No companies selected'}))
            return

        result = enrich_excel_by_companies(
            Path(excel_path),
            companies,
            cache,
            progress_path=Path(cache_path).parent / "enrich_progress.json"
        )
        self.progress_queue.put(('enrich_complete', result))

    self.current_thread = threading.Thread(target=enrich_thread, daemon=True)
    self.current_thread.start()
```

**Cambios en main.py:**

1. Cambiar checkbox por botón:
```python
# Antes (checkbox):
self.enrich_contacts_var = tk.BooleanVar(value=False)
ttk.Checkbutton(control_frame, text=" Enriquecer contactos...", variable=self.enrich_contacts_var).pack(anchor='w')

# Después (botón):
self.enrich_btn = ttk.Button(
    control_frame,
    text="🔄 Enriquecer Contactos",
    command=self._start_enrichment
)
self.enrich_btn.pack(side='left', padx=(0, 5))
```

2. Agregar método `_start_enrichment`:
```python
def _start_enrichment(self):
    """Inicia enrichment de contactos."""
    excel_path = self.scraper_output_dir.get() / "gal" / "gal_export.xlsx"
    cache_path = self.scraper_output_dir.get() / "gal" / "cache.json"

    if not excel_path.exists():
        messagebox.showwarning("Error", "Primero ejecuta Extracción GAL")
        return

    if not cache_path.exists():
        messagebox.showwarning("Error", "No se encontró cache.json")
        return

    self.service.start_enrichment(str(excel_path), str(cache_path))
    self._add_scraper_log("🔄 Iniciando enrichment...")
```

3. Agregar handler `_handle_enrich_complete` similar a `_handle_gal_complete`

4. En `_check_progress`, agregar:
```python
elif item_type == 'enrich_complete':
    self._handle_enrich_complete(data)
elif item_type == 'enrich_error':
    self._handle_enrich_error(data)
```

- [ ] **Step 1: Modificar service.py - agregar start_enrichment**

- [ ] **Step 2: Modificar main.py - quitar checkbox, agregar botón y métodos**

- [ ] **Step 3: Correr tests**

Run: `pytest tests/ -v --tb=short`

- [ ] **Step 4: Commit**

---

## Task 4: Conectar scraper GAL existente con gal_exporter

**Files:**
- Modify: `src/verificacion_correo/core/gal_scraper.py` (agregar export a Excel)

**Cambios en gal_scraper.py:**
Al final de `scrape_gal`, después de guardar JSON/CSV, agregar:
```python
# Exportar a Excel si se pide
if output_excel:
    from .gal_exporter import flatten_contact_to_dict, save_to_excel
    flattened = [flatten_contact_to_dict(p) for p in all_people]
    save_to_excel(flattened, Path(output_excel), Path(output_json))
    stats["excel"] = str(output_excel)
```

**Interfaces:**
- Consumes: `output_excel: Optional[str]` param en `scrape_gal`
- Produces: Excel con 2 hojas en `output_dir`

- [ ] **Step 1: Agregar parámetro output_excel a scrape_gal y lógica de export**

- [ ] **Step 2: Modificar GUI para pasar output_excel al scraper**

En `_start_scraper` de main.py:
```python
output_dir = Path(self.scraper_output_dir.get()) / "gal"
excel_path = output_dir / "gal_export.xlsx"

self.service.start_gal_scraping(
    output_dir=str(output_dir),
    max_contacts=max_contacts,
    output_excel=str(excel_path),  # NUEVO
    ...
)
```

- [ ] **Step 3: Tests**

- [ ] **Step 4: Commit**

---

## Verificación

1. `pytest tests/ -v` → todos pasan
2. Verificar que Excel tiene 2 hojas
3. Verificar que enrichment solo llena vacíos
4. Verificar que progress se guarda y permite reanudar
