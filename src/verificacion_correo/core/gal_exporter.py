"""GAL Excel exporter con 2 hojas: Contactos + Compañías."""

from pathlib import Path
from typing import List, Dict, Any, Optional
import json
from openpyxl import Workbook

CONTACT_FIELDS = ['nombre', 'email', 'empresa', 'telefono', 'departamento', 'oficina', 'direccion']


def flatten_contact_to_dict(persona: dict) -> dict:
    """Convierte persona del GAL a dict plano extrayendo todos los campos disponibles."""
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
            parts = [
                val.get("Street") or "",
                val.get("City") or "",
                val.get("State") or "",
                val.get("PostalCode") or "",
                val.get("Country") or "",
            ]
            parts = [p for p in parts if p.strip()]
            address = ", ".join(parts) if parts else ""
        elif isinstance(val, str):
            address = val

    return {
        'nombre': name,
        'email': email,
        'empresa': company,
        'telefono': phone,
        'departamento': department,
        'oficina': office,
        'direccion': address,
    }


def extract_companies_from_contacts(contacts: List[dict]) -> List[str]:
    """Extrae lista única de compañías de contactos."""
    companies = set()
    for c in contacts:
        company = c.get('empresa', '').strip()
        if company:
            companies.add(company)
    return sorted(companies)


def _auto_width(ws):
    """Ajusta ancho de columnas automáticamente."""
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_length + 2, 40)


def save_to_excel(contacts: List[dict], output_path: Path, cache_path: Optional[Path] = None):
    """Guarda contactos en Excel de 2 hojas.

    Sheet1: Contactos con todos los campos
    Sheet2: Compañías con checkbox Enrich
    """
    from openpyxl.styles import Font

    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Contactos"
    ws1.append(CONTACT_FIELDS)
    for cell in ws1[1]:
        cell.font = Font(bold=True)
    for contact in contacts:
        row = [contact.get(f, '') for f in CONTACT_FIELDS]
        ws1.append(row)
    _auto_width(ws1)

    ws2 = wb.create_sheet("Compañías")
    ws2.append(['Compañía', 'Enrich'])
    for cell in ws2[1]:
        cell.font = Font(bold=True)
    companies = extract_companies_from_contacts(contacts)
    for company in companies:
        ws2.append([company, ''])
    _auto_width(ws2)

    wb.save(output_path)

    if cache_path:
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(contacts, f, ensure_ascii=False, indent=2)


def load_gal_cache(cache_path: Path) -> List[dict]:
    """Carga GAL cache desde JSON."""
    with open(cache_path, 'r', encoding='utf-8') as f:
        return json.load(f)
