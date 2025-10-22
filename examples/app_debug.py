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
    name_match = NAME_RE.search(popup_text)
    name_specific = name_match.group(1).strip() if name_match else None

    # SIP
    sip_match = SIP_RE.search(popup_text)
    sip_specific = sip_match.group(0) if sip_match else None

    # Teléfono: buscar números de 9 dígitos
    phone_specific = None
    for line in popup_text.split('\n'):
        if 'sip:' not in line.lower():
            phone_match = re.search(r'\b\d{9}\b', line)
            if phone_match:
                phone_specific = phone_match.group(0)
                break

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

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state="state.json")
        page = context.new_page()

        page.goto(PAGE_URL)
        page.wait_for_load_state("networkidle")

        # Hacer clic en "Nuevo mensaje"
        page.click('button[title="Escribir un mensaje nuevo (N)"]')
        page.wait_for_timeout(1000)

        # Llenar el campo "Para"
        input_box = page.get_by_role("textbox", name="Para")
        input_box.fill(CORREO)
        page.wait_for_timeout(3000)
        input_box.blur()

        print("\n=== DEBUGGING MODE ===")
        print("INSTRUCCIONES:")
        print("1. Haz clic en uno de los tokens de email (ASP164@MADRID.ORG o AGM564@MADRID.ORG)")
        print("2. Espera a que aparezca el popup de la tarjeta de contacto")
        print("3. Presiona el botón 'Resume' en la barra de herramientas del navegador")
        print("=====================\n")

        # Pausar para que el usuario pueda hacer clic manualmente
        page.pause()

        # Intentar extraer información del popup
        print("\nExtrayendo información del popup...")
        result = popup_info(page)

        if result:
            print("\n=== INFORMACIÓN EXTRAÍDA ===")
            for key, value in result.items():
                if value:
                    print(f"{key}: {value}")
            print("============================\n")
        else:
            print("No se pudo extraer información del popup\n")

        # Cerrar el popup
        page.keyboard.press("Escape")
        page.wait_for_timeout(1000)

        # Pausar de nuevo para verificar
        print("\n¿Quieres procesar otro token? Si es así, haz clic y presiona Resume")
        page.pause()

        # Cerrar el mensaje
        try:
            page.click('button[aria-label="Descartar"]', timeout=2000)
            page.wait_for_timeout(2000)
            page.click('button[aria-label="Descartar"]', timeout=2000)
        except:
            pass

        context.close()
        browser.close()

if __name__ == "__main__":
    main()
