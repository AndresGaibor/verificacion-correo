"""GAL enrichment selectivo por compañía usando GetPersona API.

Lee directamente de Sheet1 del Excel (Contactos) y actualiza la misma fila.
No usa JSON cache - el Excel es la fuente de verdad.
"""

from pathlib import Path
from typing import List, Dict, Any, Optional, Callable
import json
import time
from openpyxl import load_workbook
from urllib.request import Request, urlopen
from urllib.error import HTTPError

from .api_extractor import _build_cookie_header, _get_canary, _build_headers, OWA_BASE


def _call_get_persona(persona_id: str, cookie_str: str, canary: str) -> Optional[dict]:
    """Llama GetPersona API y retorna dict con datos enriquecidos."""
    payload = {
        "__type": "GetPersonaJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2013",
        },
        "Body": {
            "__type": "GetPersonaRequest:#Exchange",
            "PersonaId": {
                "__type": "ItemId:#Exchange",
                "Id": persona_id,
            },
            "PersonaShape": {
                "__type": "PersonaResponseShape:#Exchange",
                "BaseShape": "Default",
            },
        },
    }

    body_bytes = json.dumps(payload).encode("utf-8")
    req = Request(
        f"{OWA_BASE}/owa/service.svc?action=GetPersona",
        data=body_bytes,
        headers=_build_headers(canary, cookie_str, "GetPersona"),
    )

    try:
        resp = urlopen(req, timeout=60)
        data = json.loads(resp.read().decode("utf-8"))
        persona = data.get("Body", {}).get("Persona")
        if not persona:
            return None

        phone = ""
        phone_items = persona.get("BusinessPhoneNumbersArray") or []
        if phone_items and isinstance(phone_items[0], dict):
            val = phone_items[0].get("Value") or {}
            if isinstance(val, dict):
                phone = val.get("Number") or val.get("NormalizedNumber") or ""

        address = ""
        addr_items = persona.get("BusinessAddressesArray") or []
        if addr_items and isinstance(addr_items[0], dict):
            val = addr_items[0].get("Value") or {}
            if isinstance(val, dict):
                parts = [
                    val.get("Street") or "",
                    val.get("City") or "",
                    val.get("PostalCode") or "",
                    val.get("State") or "",
                ]
                parts = [p for p in parts if p.strip()]
                address = ", ".join(parts)

        return {
            "telefono": phone,
            "departamento": persona.get("Department") or "",
            "oficina": "",
            "direccion": address,
        }
    except HTTPError as e:
        if e.code == 307:
            raise Exception("Session expired during GetPersona")
        return None
    except Exception:
        return None


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
       b. Para cada match con persona_id y teléfono vacío -> llama GetPersona API
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

    headers = [cell.value for cell in ws1[1]]
    persona_id_idx = headers.index('persona_id') + 1
    telefono_idx = headers.index('telefono') + 1
    depto_idx = headers.index('departamento') + 1
    oficina_idx = headers.index('oficina') + 1
    direccion_idx = headers.index('direccion') + 1
    empresa_idx = headers.index('empresa') + 1

    empresa_to_rows: Dict[str, List[int]] = {}
    for row_idx in range(2, ws1.max_row + 1):
        empresa = ws1.cell(row_idx, empresa_idx).value
        if empresa:
            empresa_to_rows.setdefault(str(empresa), []).append(row_idx)

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

            telefono_actual = ws1.cell(row_idx, telefono_idx).value
            if telefono_actual:
                continue

            try:
                enriched = _call_get_persona(str(persona_id), cookie_str, canary)
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
            companies.append(str(company))
    return companies
