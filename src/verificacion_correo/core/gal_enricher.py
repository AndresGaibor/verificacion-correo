"""GAL enrichment selectivo por compañía desde cache local."""

from pathlib import Path
from typing import List, Dict, Any, Set, Optional, Callable
import json
import time
from openpyxl import load_workbook


class EnrichProgress:
    """Trackea progreso de enrichment para reanudación."""

    def __init__(self, path: Path):
        self.path = path
        self.data: Dict[str, Any] = {
            'companies_done': [],
            'contacts_enriched': 0,
            'offset': 0,
        }

    def save(self):
        with open(self.path, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

    def load(self) -> bool:
        if self.path.exists():
            with open(self.path, 'r', encoding='utf-8') as f:
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
        companies_set -= set(progress.data.get('companies_done', []))

    wb = load_workbook(excel_path)
    ws1 = wb["Contactos"]
    ws2 = wb["Compañías"]

    # Construir índice de cache por (email_lower, empresa)
    cache_index: Dict[tuple, dict] = {}
    for c in cache:
        key = (c.get('email', '').lower(), c.get('empresa', ''))
        cache_index[key] = c

    # Headers de Sheet1
    headers = [cell.value for cell in ws1[1]]
    telefono_idx = headers.index('telefono') + 1
    depto_idx = headers.index('departamento') + 1
    oficina_idx = headers.index('oficina') + 1
    direccion_idx = headers.index('direccion') + 1
    empresa_idx = headers.index('empresa') + 1
    email_idx = headers.index('email') + 1

    total_enriched = 0
    companies_done: List[str] = []

    # Iterar Sheet2 para encontrar empresas marcadas
    for row_idx in range(2, ws2.max_row + 1):
        company = ws2.cell(row_idx, 1).value
        enrich_mark = ws2.cell(row_idx, 2).value

        enrich_str = str(enrich_mark).strip().upper() if enrich_mark else ''
        if not company or enrich_str != 'X':
            continue
        if company not in companies_set:
            continue

        # Buscar contactos en Sheet1 que matcheen esta compañía
        for data_row_idx in range(2, ws1.max_row + 1):
            empresa_cell = ws1.cell(data_row_idx, empresa_idx).value
            if empresa_cell != company:
                continue

            email_cell = ws1.cell(data_row_idx, email_idx).value
            cache_key = (str(email_cell).lower() if email_cell else '', company)
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


def get_companies_to_enrich_from_excel(excel_path: Path) -> List[str]:
    """Lee Sheet2 y retorna lista de compañías marcadas con X."""
    wb = load_workbook(excel_path)
    ws2 = wb["Compañías"]
    companies = []
    for row_idx in range(2, ws2.max_row + 1):
        company = ws2.cell(row_idx, 1).value
        enrich_mark = ws2.cell(row_idx, 2).value
        enrich_str = str(enrich_mark).strip().upper() if enrich_mark else ''
        if company and enrich_str == 'X':
            companies.append(company)
    return companies
