from playwright.sync_api import Page, sync_playwright, TimeoutError as PWTimeout
import re

PAGE_URL = "https://correoweb.madrid.org/owa/#path=/mail"
CORREO = "ASP164@MADRID.ORG;AGM564@MADRID.ORG"

EMAIL_RE = re.compile(r'[\w.+-]+@[\w.-]+\.[a-z]{2,}', re.I)
PHONE_RE = re.compile(r'\b\d{6,}\b')
POSTAL_ADDR_RE = re.compile(r'\d{5}\s+[A-ZÁÉÍÓÚÑ\-\s]+', re.I)
SIP_RE = re.compile(r'sip:[\w.+-]+@[\w.-]+', re.I)
NAME_RE = re.compile(r'([A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúñ\-\.\s]+,\s*[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúñ\-\s]+)')

def extract_from_popup_text(text: str) -> dict:
    """Busca por regex en el texto del popup para devolver campos posibles."""
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
            if line_stripped not in ['CONTACTO', 'NOTAS', 'ORGANIZACIÓN', 'CALENDARIO', 'TRABAJO',
                                     'ENVIAR CORREO ELECTRÓNICO', 'DIRECCIÓN PROFESIONAL']:
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
    """Extrae información del popup de tarjeta de contacto."""
    # Esperar a que el popup aparezca
    try:
        popup = page.locator('div._pe_Y[ispopup="1"]').first
        popup.wait_for(state="visible", timeout=5000)
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

    # SIP
    sip_match = SIP_RE.search(popup_text)
    sip_specific = sip_match.group(0) if sip_match else None

    # Teléfono: buscar números de 6+ dígitos (flexibilidad para diferentes formatos)
    phone_specific = None
    for line in popup_text.split('\n'):
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

def click_email_token(page: Page, email: str) -> bool:
    """
    Intenta hacer clic en un token de email usando diferentes estrategias.
    Retorna True si el popup aparece, False en caso contrario.
    """
    print(f"  Buscando token: {email}")

    # Estrategia 1: Buscar el span que contiene el email y subir al contenedor padre
    try:
        # Buscar spans que contengan exactamente el email (case-insensitive)
        email_span = page.locator(f'span:has-text("{email}")').filter(has_text=re.compile(f'^{re.escape(email)}$', re.I))

        if email_span.count() > 0:
            print(f"  Encontrado {email_span.count()} span(s) con el email")

            # Intentar con el primer match
            first_span = email_span.first

            # Intentar hacer clic directamente en el span
            print("  Estrategia 1: Click directo en el span del email")
            first_span.click(timeout=3000)
            page.wait_for_timeout(2000)

            # Verificar si apareció el popup
            if page.locator('div._pe_Y[ispopup="1"]').first.is_visible():
                print("  ✓ Popup abierto con click directo")
                return True

            # Si no funcionó, intentar con el contenedor padre
            print("  Estrategia 2: Click en el contenedor padre")
            # Subir al contenedor del token (generalmente un div)
            parent = first_span.locator('xpath=ancestor::div[contains(@class, "_pe_")]').first
            if parent.count() > 0:
                parent.click(timeout=3000)
                page.wait_for_timeout(2000)

                if page.locator('div._pe_Y[ispopup="1"]').first.is_visible():
                    print("  ✓ Popup abierto con click en contenedor padre")
                    return True

            # Estrategia 3: Hacer doble click
            print("  Estrategia 3: Doble click en el span")
            first_span.dblclick(timeout=3000)
            page.wait_for_timeout(2000)

            if page.locator('div._pe_Y[ispopup="1"]').first.is_visible():
                print("  ✓ Popup abierto con doble click")
                return True

    except Exception as e:
        print(f"  Error al hacer clic en el token: {e}")

    return False

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state="state.json")
        page = context.new_page()

        page.goto(PAGE_URL)
        page.wait_for_load_state("networkidle")

        # Hacer clic en "Nuevo mensaje"
        print("Abriendo nuevo mensaje...")
        page.click('button[title="Escribir un mensaje nuevo (N)"]')
        page.wait_for_timeout(1000)

        # Llenar el campo "Para"
        print(f"Llenando campo Para con: {CORREO}")
        input_box = page.get_by_role("textbox", name="Para")
        input_box.fill(CORREO)
        page.wait_for_timeout(3000)
        input_box.blur()

        # Parsear emails
        emails = [e.strip() for e in CORREO.split(";") if e.strip()]
        print(f"\nEmails a procesar: {emails}")

        # Procesar cada email
        results = []
        for i, email in enumerate(emails):
            print(f"\n{'='*60}")
            print(f"Procesando email {i+1}/{len(emails)}: {email}")
            print('='*60)

            # Intentar hacer clic en el token
            if click_email_token(page, email):
                # Extraer información del popup
                result = popup_info(page)

                if result:
                    print("\n  === INFORMACIÓN EXTRAÍDA ===")
                    for key, value in result.items():
                        if value:
                            print(f"  {key}: {value}")
                    print("  " + "="*28)
                    results.append({**result, "token_email": email})
                else:
                    print("  ✗ No se pudo extraer información del popup")

                # Cerrar el popup
                page.keyboard.press("Escape")
                page.wait_for_timeout(1000)
            else:
                print("  ✗ No se pudo abrir el popup para este email")
                print("  Pausando para debugging manual...")
                page.pause()

        # Resumen
        print(f"\n{'='*60}")
        print(f"RESUMEN: Procesados {len(results)}/{len(emails)} emails")
        print('='*60)

        # Cerrar el mensaje
        print("\nCerrando mensaje...")
        try:
            page.click('button[aria-label="Descartar"]', timeout=2000)
            page.wait_for_timeout(2000)
            page.click('button[aria-label="Descartar"]', timeout=2000)
        except:
            pass

        context.close()
        browser.close()

        return results

if __name__ == "__main__":
    resultados = main()

    # Mostrar resultados finales
    print("\n" + "="*60)
    print("RESULTADOS FINALES")
    print("="*60)
    for i, r in enumerate(resultados, 1):
        print(f"\n{i}. {r.get('token_email', 'N/A')}")
        print(f"   Nombre: {r.get('name', 'N/A')}")
        print(f"   Email: {r.get('email', 'N/A')}")
        print(f"   Teléfono: {r.get('phone', 'N/A')}")
        print(f"   Oficina: {r.get('office', 'N/A')}")
