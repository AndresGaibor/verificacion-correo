"""
Módulo para automatización del navegador con Playwright (versión async con ADFS recovery).

Procesa emails en lotes y guarda resultados incrementalmente al Excel.
Si la sesión ADFS/OWA expira, detecta el redirect, descarta el contexto
con cookies viejas y espera a que el usuario haga login manual.
"""

import asyncio
from playwright.async_api import async_playwright, Page, TimeoutError as PWTimeout
from config import PAGE_URL, SELECTORS, WAIT_TIMES, BROWSER_CONFIG
from contact_extractor import popup_info
from excel_writer import escribir_resultado


async def _esperar_login_manual(page: Page, url: str, timeout_segundos: int = 300) -> bool:
    """
    Espera a que el usuario complete el login manual en ADFS.

    Detecta login exitoso cuando la URL vuelve a contener correoweb.madrid.org/owa.
    Reintenta la navegación si aparece TimeoutLogout.aspx.

    Args:
        page: Página actual con la pantalla de login
        url: URL a la que reintentar ir
        timeout_segundos: Tiempo máximo de espera (default 5 minutos)

    Returns:
        True si el login fue exitoso, False si se agotó el tiempo
    """
    print("[*] Inicie sesion manualmente en la ventana del navegador.")
    print(f"[...] Esperando hasta {timeout_segundos // 60} minutos a que inicie sesion...")

    for _ in range(timeout_segundos):
        try:
            url_actual = page.url
        except Exception:
            await asyncio.sleep(1)
            continue

        if 'correoweb.madrid.org/owa' in url_actual.lower():
            print("[OK] Login detectado - sesion establecida")
            return True

        if 'timeout' in url_actual.lower() or 'logout' in url_actual.lower():
            print("[!] Detectado TimeoutLogout - reintentando navegacion...")
            try:
                await page.goto(url, wait_until="load", timeout=120000)
            except Exception as e:
                print(f"   Error reintentando: {e}")

        await asyncio.sleep(1)

    print("[!] Tiempo de espera agotado - continuando de todos modos")
    return False


async def procesar_lote(browser, context, lote, archivo_excel, numero_lote, total_lotes, page: Page | None = None):
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
        page: Página existente a reutilizar (opcional)

    Returns:
        Dict con estadísticas: {'ok': int, 'no_existe': int, 'error': int}
    """
    stats = {'ok': 0, 'no_existe': 0, 'error': 0}

    print(f"\nProcesando lote {numero_lote}/{total_lotes}...")

    if page is None:
        if context.pages:
            page = context.pages[0]
        else:
            page = await context.new_page()

    try:
        if page.url != PAGE_URL and not page.url.startswith(PAGE_URL.split('#')[0]):
            await page.goto(PAGE_URL)
            await page.wait_for_load_state("load", timeout=120000)

        await page.click(SELECTORS['new_message_btn'])
        await page.wait_for_timeout(WAIT_TIMES['after_new_message'])

        input_box = page.get_by_role(SELECTORS['to_field_role'], name=SELECTORS['to_field_name'])
        emails_str = ";".join([item['email'] for item in lote])
        await input_box.fill(emails_str)
        await page.wait_for_timeout(WAIT_TIMES['after_fill_to'])
        await input_box.blur()
        await page.wait_for_timeout(WAIT_TIMES['after_blur'])

        for idx, item in enumerate(lote, 1):
            email = item['email']
            fila = item['fila']

            print(f"  [{idx}/{len(lote)}] {email} ...", end=' ')

            try:
                # OWA inserta el email en un span con texto del nombre y el email puede
                # estar en atributos title, aria-label, o en el src de la foto cercana.
                # Estrategia: revisar inner_text, title y aria-label de cada span.
                email_normalized = email.strip().lower()
                email_at = email_normalized.replace('@', '%40')

                # 1) Buscar span cuyo title/aria-label/inner_text contenga el email
                spans = await page.locator('span').all()
                email_span = None
                for cand in spans:
                    try:
                        inner = (await cand.inner_text() or '').strip().lower()
                        title = ((await cand.get_attribute('title')) or '').strip().lower()
                        aria = ((await cand.get_attribute('aria-label')) or '').strip().lower()
                        if email_normalized in (inner, title, aria):
                            email_span = cand
                            break
                    except Exception:
                        continue

                # 2) Si no, buscar imagen con src que contenga el email codificado
                if email_span is None:
                    imgs = await page.locator(f'img[src*="{email_at}"]').all()
                    if imgs:
                        # Subir al ancestro que tenga clase _rw_k (estructura del token)
                        email_span = imgs[0]
                        try:
                            email_span = await email_span.evaluate_handle(
                                "el => el.closest('._rw_k') || el.closest('span') || el"
                            )
                        except Exception:
                            pass

                if email_span is None:
                    print(f"[X] Token no encontrado (buscado: {email})")
                    escribir_resultado(archivo_excel, fila, "ERROR")
                    stats['error'] += 1
                    continue

                await email_span.as_element().click(timeout=3000) if hasattr(email_span, 'as_element') else await email_span.click(timeout=3000)
                await page.wait_for_timeout(WAIT_TIMES['after_click_token'])

                popup_locator = page.locator(SELECTORS['popup']).first
                try:
                    visible = await popup_locator.is_visible()
                except Exception:
                    visible = False

                if not visible:
                    print("[X] Popup no se abrio")
                    escribir_resultado(archivo_excel, fila, "ERROR")
                    stats['error'] += 1
                    continue

                resultado = await popup_info(page)

                if resultado and resultado.get('sip') and resultado.get('sip').strip():
                    escribir_resultado(archivo_excel, fila, "OK", resultado)
                    stats['ok'] += 1
                    print("[OK] OK")
                elif resultado and any(resultado.values()):
                    escribir_resultado(archivo_excel, fila, "NO EXISTE")
                    stats['no_existe'] += 1
                    print("[NO] NO EXISTE (sin SIP)")
                else:
                    escribir_resultado(archivo_excel, fila, "NO EXISTE")
                    stats['no_existe'] += 1
                    print("[NO] NO EXISTE")

                await page.keyboard.press("Escape")
                await page.wait_for_timeout(WAIT_TIMES['after_close_popup'])

            except Exception as e:
                print(f"✗ ERROR ({str(e)[:50]})")
                escribir_resultado(archivo_excel, fila, "ERROR")
                stats['error'] += 1

        await _cerrar_mensaje(page)

        print(f"Lote {numero_lote} completado: {stats['ok']} OK, {stats['no_existe']} NO EXISTE, {stats['error']} ERROR")

    except Exception as e:
        print(f"ERROR procesando lote {numero_lote}: {e}")
        for item in lote:
            try:
                escribir_resultado(archivo_excel, item['fila'], "ERROR")
                stats['error'] += 1
            except Exception:
                pass

    return stats


async def procesar_todos_los_lotes(lotes_info, archivo_excel):
    """
    Procesa todos los lotes de emails secuencialmente con manejo de ADFS.

    Si detecta que la sesión expiró (redirect a login), cierra el contexto
    con cookies viejas, crea uno nuevo y espera login manual.

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

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=BROWSER_CONFIG['headless'])
        context = await browser.new_context(storage_state=BROWSER_CONFIG['session_file'])
        page = await context.new_page()

        try:
            await page.goto(PAGE_URL, wait_until="load", timeout=120000)

            if ('login' in page.url.lower() or 'signin' in page.url.lower()
                    or 'adfs' in page.url.lower() or 'sts.' in page.url.lower()):
                print("[!] Sesion expirada - redirigido a pagina de login")
                print("[*] Cerrando contexto con sesion vieja y creando uno nuevo...")

                await page.close()
                await context.close()

                context = await browser.new_context()
                page = await context.new_page()
                await page.goto(PAGE_URL, wait_until="load", timeout=120000)

                if not await _esperar_login_manual(page, PAGE_URL):
                    print("[X] No se pudo restablecer la sesion")
                    await context.close()
                    await browser.close()
                    return total_stats

                from pathlib import Path
                import json
                state_path = Path(BROWSER_CONFIG['session_file']).resolve()
                storage = await context.storage_state()
                with open(state_path, 'w') as f:
                    json.dump(storage, f, indent=2)
                print(f"[+] Nueva sesion guardada en {state_path}")

            for idx, lote in enumerate(lotes, 1):
                lote_stats = await procesar_lote(
                    browser, context, lote, archivo_excel, idx, len(lotes), page
                )
                total_stats['ok'] += lote_stats['ok']
                total_stats['no_existe'] += lote_stats['no_existe']
                total_stats['error'] += lote_stats['error']

        finally:
            try:
                await context.close()
            except Exception:
                pass
            try:
                await browser.close()
            except Exception:
                pass

    return total_stats


async def _cerrar_mensaje(page: Page):
    """
    Cierra el mensaje sin guardar (descarta el borrador).

    Args:
        page: Objeto Page async de Playwright
    """
    try:
        await page.wait_for_timeout(WAIT_TIMES['before_discard'])
        await page.click(SELECTORS['discard_btn'], timeout=2000)
        await page.wait_for_timeout(1000)
        await page.click(SELECTORS['discard_btn'], timeout=2000)
    except PWTimeout:
        pass
    except Exception:
        pass
