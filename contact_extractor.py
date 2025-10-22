"""
Módulo para extraer información de contacto desde popups de OWA.
"""

import re
from playwright.sync_api import Page, TimeoutError as PWTimeout
from config import EMAIL_RE, PHONE_RE, POSTAL_ADDR_RE, SIP_RE, NAME_RE, WAIT_TIMES, SELECTORS


def extract_from_popup_text(text: str) -> dict:
    """
    Busca por regex en el texto del popup para devolver campos posibles.

    Args:
        text: Texto completo del popup

    Returns:
        Diccionario con los campos extraídos
    """
    email = EMAIL_RE.search(text)
    sip = SIP_RE.search(text)
    phone = PHONE_RE.search(text)
    postal = POSTAL_ADDR_RE.search(text)
    name = NAME_RE.search(text)

    # Cargo/oficina: heurística — busca líneas en mayúsculas (por el HTML mostrado)
    office = None
    for line in text.splitlines():
        line_stripped = line.strip()
        if line_stripped and line_stripped.isupper() and len(line_stripped) > 3:
            # evitar capturar textos genéricos; tomar la primera línea en mayúsculas que parezca oficina
            office = line_stripped
            break

    return {
        "name": name.group(1).strip() if name else None,
        "email": email.group(0).strip() if email else None,
        "sip": sip.group(0).strip() if sip else None,
        "phone": phone.group(0).strip() if phone else None,
        "postal_or_address": postal.group(0).strip() if postal else None,
        "office_or_job": office
    }


def popup_info(page: Page) -> dict:
    """
    Extrae información del popup de tarjeta de contacto.

    Args:
        page: Objeto Page de Playwright

    Returns:
        Diccionario con la información extraída del popup o None si no se encuentra
    """
    # Esperar a que el popup aparezca
    try:
        popup = page.locator(SELECTORS['popup']).first
        popup.wait_for(state="visible", timeout=WAIT_TIMES['popup_visible'])
    except PWTimeout:
        print("No se encontró el popup de tarjeta de contacto")
        return None

    # Extraer todo el texto visible del popup
    popup_text = popup.inner_text()

    # Buscar emails que NO sean ASP/AGM (emails genéricos)
    email_specific = None
    all_emails = EMAIL_RE.findall(popup_text)
    for email in all_emails:
        # Filtrar emails genéricos (ASP, AGM, etc.)
        if not re.match(r'^(ASP|AGM|AEM|ADM)\d+@', email, re.I):
            email_specific = email
            break

    # Buscar nombre en formato "APELLIDO, NOMBRE"
    # Primero intentar con regex
    name_match = NAME_RE.search(popup_text)
    name_specific = name_match.group(1).strip() if name_match else None

    # Si no funcionó, buscar por el heading del popup que suele tener el nombre
    if not name_specific:
        lines = popup_text.split('\n')
        for line in lines[:10]:  # Buscar en las primeras 10 líneas
            line = line.strip()
            # Buscar línea que tenga coma y letras mayúsculas (formato nombre)
            if ',' in line and len(line) > 5:
                # Verificar que no sea un label o texto genérico
                if line not in ['CONTACTO', 'NOTAS', 'ORGANIZACIÓN'] and not line.startswith('C/'):
                    # Verificar que tenga formato de nombre (al menos una coma y letras)
                    if re.match(r'^[A-ZÁÉÍÓÚÑ\s]+,\s*[A-ZÁÉÍÓÚÑ\s]+$', line, re.I):
                        name_specific = line
                        break

    # SIP - Extraer y limpiar
    sip_match = SIP_RE.search(popup_text)
    sip_specific = sip_match.group(0).strip() if sip_match else None

    # Validación adicional: asegurar que el SIP tenga formato válido
    if sip_specific and not re.match(r'^sip:[\w.+-]+@[\w.-]+\.[a-z]{2,}$', sip_specific, re.I):
        # Si no tiene extensión de dominio, intentar corregir
        sip_specific = None

    # Teléfono: buscar números de 6+ dígitos (flexibilidad para diferentes formatos)
    phone_specific = None
    lines = popup_text.split('\n')
    for line in lines:
        if 'sip:' not in line.lower() and not re.search(r'\d{5}\s+[A-Z]', line):  # No códigos postales
            # Primero intentar 9 dígitos
            phone_match = re.search(r'\b\d{9}\b', line)
            if phone_match:
                phone_specific = phone_match.group(0)
                break
            # Si no, intentar 6-8 dígitos
            if not phone_specific:
                phone_match = re.search(r'\b\d{6,8}\b', line)
                if phone_match and 'Trabajo' in popup_text[:popup_text.find(phone_match.group(0)) + 100]:
                    phone_specific = phone_match.group(0)

    # Dirección
    addr_match = re.search(r'C/\s+[A-ZÁÉÍÓÚÑ\s,]+\d+\s+\d{5}\s+[A-ZÁÉÍÓÚÑ\-\s]+', popup_text, re.I)
    addr_specific = addr_match.group(0).strip() if addr_match else None

    # Información del trabajo
    dept_match = re.search(r'Departamento:\s*([^\n]+)', popup_text)
    department = dept_match.group(1).strip() if dept_match else None

    comp_match = re.search(r'Compañía:\s*([^\n]+)', popup_text)
    company = comp_match.group(1).strip() if comp_match else None

    office_match = re.search(r'Oficina:\s*([^\n]+)', popup_text)
    office_location = office_match.group(1).strip() if office_match else None

    # Consolidar usando también la extracción regex de fallback
    consolidated = extract_from_popup_text(popup_text)

    result = {
        "name": name_specific or consolidated["name"],
        "email": email_specific or consolidated["email"],
        "sip": sip_specific or consolidated["sip"],
        "phone": phone_specific or consolidated["phone"],
        "address": addr_specific or consolidated["postal_or_address"],
        "office": department,
        "department": department,
        "company": company,
        "office_location": office_location,
    }

    return result
