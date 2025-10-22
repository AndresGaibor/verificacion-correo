"""
Módulo para automatización del navegador con Playwright.
Procesa emails en lotes y guarda resultados incrementalmente al Excel.
"""

import re
from playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeout
from config import PAGE_URL, SELECTORS, WAIT_TIMES, BROWSER_CONFIG
from contact_extractor import popup_info
from excel_writer import escribir_resultado


def procesar_lote(browser, context, lote, archivo_excel, numero_lote, total_lotes):
    """
    Procesa un lote de emails (máx 10).
    Abre un mensaje, agrega los emails, extrae información de cada uno,
    y guarda cada resultado inmediatamente al Excel.

    Args:
        browser: Instancia del navegador Playwright
        context: Contexto del navegador
        lote: Lista de dicts {'email': str, 'fila': int}
        archivo_excel: Ruta al archivo Excel
        numero_lote: Número de este lote (para mostrar progreso)
        total_lotes: Total de lotes a procesar

    Returns:
        Dict con estadísticas: {'ok': int, 'no_existe': int, 'error': int}
    """
    stats = {'ok': 0, 'no_existe': 0, 'error': 0}

    print(f"\nProcesando lote {numero_lote}/{total_lotes}...")

    # Si no hay página activa, crear una nueva
    if context.pages:
        page = context.pages[0]
    else:
        page = context.new_page()

    try:
        # Navegar a OWA si no estamos ahí
        if page.url != PAGE_URL and not page.url.startswith(PAGE_URL.split('#')[0]):
            page.goto(PAGE_URL)
            page.wait_for_load_state("networkidle")

        # Abrir nuevo mensaje
        page.click(SELECTORS['new_message_btn'])
        page.wait_for_timeout(WAIT_TIMES['after_new_message'])

        # Agregar todos los emails del lote al campo "Para"
        input_box = page.get_by_role(SELECTORS['to_field_role'], name=SELECTORS['to_field_name'])
        emails_str = ";".join([item['email'] for item in lote])
        input_box.fill(emails_str)
        page.wait_for_timeout(WAIT_TIMES['after_fill_to'])
        input_box.blur()
        page.wait_for_timeout(WAIT_TIMES['after_blur'])

        # Procesar cada email individualmente
        for idx, item in enumerate(lote, 1):
            email = item['email']
            fila = item['fila']

            print(f"  [{idx}/{len(lote)}] {email} ...", end=' ')

            try:
                # Buscar el token del email
                email_span = page.locator(f'span:has-text("{email}")').filter(
                    has_text=re.compile(f'^{re.escape(email)}$', re.I)
                )

                if email_span.count() == 0:
                    print("✗ Token no encontrado")
                    escribir_resultado(archivo_excel, fila, "ERROR")
                    stats['error'] += 1
                    continue

                # Hacer clic en el token
                email_span.first.click(timeout=3000)
                page.wait_for_timeout(WAIT_TIMES['after_click_token'])

                # Verificar que el popup esté abierto
                popup_locator = page.locator(SELECTORS['popup']).first
                if not popup_locator.is_visible():
                    print("✗ Popup no se abrió")
                    escribir_resultado(archivo_excel, fila, "ERROR")
                    stats['error'] += 1
                    continue

                # Extraer información del popup
                resultado = popup_info(page)

                # Validación específica: requerir SIP para considerar válido
                if resultado and resultado.get('sip') and resultado.get('sip').strip():
                    # Hay SIP válido, guardar como OK
                    escribir_resultado(archivo_excel, fila, "OK", resultado)
                    stats['ok'] += 1
                    print("✓ OK")
                elif resultado and any(resultado.values()):
                    # Hay datos pero no SIP, marcar como NO EXISTE (inválido)
                    escribir_resultado(archivo_excel, fila, "NO EXISTE")
                    stats['no_existe'] += 1
                    print("✗ NO EXISTE (sin SIP)")
                else:
                    # No hay datos, marcar como NO EXISTE
                    escribir_resultado(archivo_excel, fila, "NO EXISTE")
                    stats['no_existe'] += 1
                    print("✗ NO EXISTE")

                # Cerrar el popup
                page.keyboard.press("Escape")
                page.wait_for_timeout(WAIT_TIMES['after_close_popup'])

            except Exception as e:
                print(f"✗ ERROR ({str(e)[:50]})")
                escribir_resultado(archivo_excel, fila, "ERROR")
                stats['error'] += 1

        # Cerrar el mensaje sin guardar
        _cerrar_mensaje(page)

        print(f"Lote {numero_lote} completado: {stats['ok']} OK, {stats['no_existe']} NO EXISTE, {stats['error']} ERROR")

    except Exception as e:
        print(f"ERROR procesando lote {numero_lote}: {e}")
        # Marcar todos los emails restantes del lote como ERROR
        for item in lote:
            try:
                escribir_resultado(archivo_excel, item['fila'], "ERROR")
                stats['error'] += 1
            except:
                pass

    return stats


def procesar_todos_los_lotes(lotes_info, archivo_excel):
    """
    Procesa todos los lotes de emails secuencialmente.

    Args:
        lotes_info: Dict con información de lotes (de leer_correos_pendientes)
        archivo_excel: Ruta al archivo Excel

    Returns:
        Dict con estadísticas finales
    """
    lotes = lotes_info['lotes']

    if not lotes:
        print("No hay emails pendientes para procesar.")
        return {'ok': 0, 'no_existe': 0, 'error': 0}

    total_stats = {'ok': 0, 'no_existe': 0, 'error': 0}

    with sync_playwright() as p:
        # Lanzar navegador con sesión guardada
        browser = p.chromium.launch(headless=BROWSER_CONFIG['headless'])
        context = browser.new_context(storage_state=BROWSER_CONFIG['session_file'])

        # Procesar cada lote
        for idx, lote in enumerate(lotes, 1):
            lote_stats = procesar_lote(
                browser,
                context,
                lote,
                archivo_excel,
                idx,
                len(lotes)
            )

            # Acumular estadísticas
            total_stats['ok'] += lote_stats['ok']
            total_stats['no_existe'] += lote_stats['no_existe']
            total_stats['error'] += lote_stats['error']

        # Cerrar navegador
        context.close()
        browser.close()

    return total_stats


def _cerrar_mensaje(page: Page):
    """
    Cierra el mensaje sin guardar (descarta el borrador).

    Args:
        page: Objeto Page de Playwright
    """
    try:
        page.wait_for_timeout(WAIT_TIMES['before_discard'])
        page.click(SELECTORS['discard_btn'], timeout=2000)
        page.wait_for_timeout(1000)
        # Confirmar descarte (segundo clic)
        page.click(SELECTORS['discard_btn'], timeout=2000)
    except PWTimeout:
        pass
    except Exception:
        pass
