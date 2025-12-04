"""
Script de debug para scraping web con sesión persistente
Ejecutar con: python debug_scraper.py
"""
import asyncio
import os
from pathlib import Path
from playwright.async_api import async_playwright, Browser, BrowserContext, Page
import logging
import traceback
from datetime import datetime
from functools import wraps

# ============================================================================
# CONFIGURACIÓN DE LOGGING AUTOMÁTICO DE ERRORES
# ============================================================================

# Crear directorio de logs si no existe
logs_dir = Path(__file__).parent / "logs"
logs_dir.mkdir(exist_ok=True)

# Configurar logger principal
logger = logging.getLogger("debug_scraper")
logger.setLevel(logging.DEBUG)

# Handler para archivo - guarda TODO (incluyendo DEBUG)
log_file = logs_dir / f"scraper_{datetime.now().strftime('%Y%m%d')}.log"
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setLevel(logging.DEBUG)

# Handler para errores críticos en archivo separado
error_file = logs_dir / f"scraper_errors_{datetime.now().strftime('%Y%m%d')}.log"
error_handler = logging.FileHandler(error_file, encoding='utf-8')
error_handler.setLevel(logging.ERROR)

# Formato detallado para logs
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
file_handler.setFormatter(formatter)
error_handler.setFormatter(formatter)

# Handler para consola (solo INFO y superior)
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_formatter = logging.Formatter('%(message)s')
console_handler.setFormatter(console_formatter)

# Agregar handlers
logger.addHandler(file_handler)
logger.addHandler(error_handler)
logger.addHandler(console_handler)

logger.info("=" * 60)
logger.info(f"🚀 Iniciando Debug Scraper - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
logger.info(f"📝 Logs guardándose en: {log_file}")
logger.info(f"❌ Errores guardándose en: {error_file}")
logger.info("=" * 60)


def log_errors(func):
    """Decorador para capturar y loguear errores automáticamente"""
    @wraps(func)
    async def async_wrapper(*args, **kwargs):
        try:
            return await func(*args, **kwargs)
        except Exception as e:
            logger.error(f"❌ Error en {func.__name__}: {str(e)}")
            logger.error(f"Traceback completo:\n{traceback.format_exc()}")
            
            # Intentar tomar screenshot si hay página disponible
            try:
                if 'page' in kwargs or (args and hasattr(args[0], 'page')):
                    page = kwargs.get('page') or (args[0].page if hasattr(args[0], 'page') else None)
                    if page:
                        screenshot_dir = Path(__file__).parent / "logs" / "screenshots"
                        screenshot_dir.mkdir(exist_ok=True)
                        screenshot_path = screenshot_dir / f"error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                        await page.screenshot(path=str(screenshot_path))
                        logger.error(f"📸 Screenshot del error guardado: {screenshot_path}")
            except Exception as screenshot_error:
                logger.error(f"No se pudo tomar screenshot: {screenshot_error}")
            
            raise
    
    @wraps(func)
    def sync_wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logger.error(f"❌ Error en {func.__name__}: {str(e)}")
            logger.error(f"Traceback completo:\n{traceback.format_exc()}")
            raise
    
    if asyncio.iscoroutinefunction(func):
        return async_wrapper
    else:
        return sync_wrapper


# Manejador global de excepciones no capturadas
def handle_exception(exc_type, exc_value, exc_traceback):
    """Captura excepciones no manejadas y las registra"""
    if issubclass(exc_type, KeyboardInterrupt):
        # Dejar que KeyboardInterrupt se maneje normalmente
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    
    logger.critical("❌ EXCEPCIÓN NO CAPTURADA", exc_info=(exc_type, exc_value, exc_traceback))
    logger.critical(f"Tipo: {exc_type.__name__}")
    logger.critical(f"Valor: {exc_value}")
    logger.critical(f"Traceback:\n{''.join(traceback.format_tb(exc_traceback))}")

import sys
sys.excepthook = handle_exception

# ============================================================================


class DebugScraper:
    def __init__(self, headless: bool = False, session_file: str = "state.json"):
        self.headless = headless
        self.browser: Browser | None = None
        self.context: BrowserContext | None = None
        self.page: Page | None = None
        self.playwright = None
        # Usar el archivo state.json del proyecto
        self.session_file = Path(session_file).resolve()
        
    @log_errors
    async def iniciar_sesion(self, url: str = "https://www.google.com"):
        """Inicializa el navegador y carga la sesión desde state.json"""
        print("🚀 Iniciando navegador...")
        
        self.playwright = await async_playwright().start()
        
        # Verificar si existe el archivo de sesión
        if self.session_file.exists():
            print(f"✅ Archivo de sesión encontrado: {self.session_file}")
        else:
            print(f"⚠️  Archivo de sesión NO encontrado: {self.session_file}")
            print("   El navegador se iniciará sin sesión guardada")
        
        # Iniciar navegador
        self.browser = await self.playwright.chromium.launch(
            headless=self.headless,
            args=[
                '--disable-blink-features=AutomationControlled',
                '--no-sandbox',
            ],
        )
        
        # Crear contexto con sesión guardada (si existe)
        context_options = {
            'viewport': {'width': 1280, 'height': 720},
        }
        
        if self.session_file.exists():
            context_options['storage_state'] = str(self.session_file)
            print("🔑 Cargando sesión desde state.json...")
        
        self.context = await self.browser.new_context(**context_options)
        
        # Crear nueva página
        self.page = await self.context.new_page()
        
        print(f"✅ Navegador iniciado (headless={self.headless})")
        
        # Navegar a URL inicial
        print(f"🌐 Navegando a: {url}")
        await self.page.goto(url, wait_until="networkidle")
        print("✅ Página cargada")
        
        return self.page
    
    async def listar_directorio(self):
        """Lista elementos en la página actual (similar a listar directorio)"""
        print("\n📂 LISTANDO ELEMENTOS EN LA PÁGINA:")
        print("=" * 60)
        
        # Listar links
        links = await self.page.locator('a').all()
        print(f"\n🔗 Links encontrados: {len(links)}")
        for i, link in enumerate(links[:10], 1):  # Primeros 10
            try:
                text = await link.inner_text()
                href = await link.get_attribute('href')
                print(f"  {i}. {text[:50]} -> {href}")
            except:
                pass
        
        # Listar imágenes
        images = await self.page.locator('img').all()
        print(f"\n🖼️  Imágenes encontradas: {len(images)}")
        for i, img in enumerate(images[:5], 1):  # Primeras 5
            try:
                src = await img.get_attribute('src')
                alt = await img.get_attribute('alt')
                print(f"  {i}. {alt} -> {src}")
            except:
                pass
        
        # Listar botones
        buttons = await self.page.locator('button').all()
        print(f"\n🔘 Botones encontrados: {len(buttons)}")
        for i, btn in enumerate(buttons[:10], 1):
            try:
                text = await btn.inner_text()
                print(f"  {i}. {text[:50]}")
            except:
                pass
        
        print("=" * 60)
    
    async def extraer_datos(self, selector: str):
        """Extrae datos de un selector específico"""
        print(f"\n🔍 Extrayendo datos de: {selector}")
        elementos = await self.page.locator(selector).all()
        
        datos = []
        for elemento in elementos:
            try:
                texto = await elemento.inner_text()
                datos.append(texto)
            except:
                pass
        
        print(f"✅ Extraídos {len(datos)} elementos")
        return datos
    
    async def esperar_elemento(self, selector: str, timeout: int = 5000):
        """Espera a que un elemento esté visible"""
        print(f"⏳ Esperando elemento: {selector}")
        try:
            await self.page.wait_for_selector(selector, timeout=timeout)
            print("✅ Elemento encontrado")
            return True
        except Exception as e:
            print(f"❌ Error esperando elemento: {e}")
            return False
    
    async def click_elemento(self, selector: str):
        """Hace click en un elemento"""
        print(f"🖱️  Haciendo click en: {selector}")
        try:
            await self.page.click(selector)
            print("✅ Click exitoso")
            return True
        except Exception as e:
            print(f"❌ Error al hacer click: {e}")
            return False
    
    async def llenar_formulario(self, campos: dict):
        """Llena un formulario con los datos proporcionados"""
        print("📝 Llenando formulario...")
        for selector, valor in campos.items():
            try:
                await self.page.fill(selector, valor)
                print(f"  ✅ {selector} = {valor}")
            except Exception as e:
                print(f"  ❌ Error en {selector}: {e}")
    
    async def tomar_screenshot(self, nombre: str = "screenshot.png"):
        """Toma una captura de pantalla"""
        ruta = Path(__file__).parent / nombre
        await self.page.screenshot(path=str(ruta))
        print(f"📸 Screenshot guardado: {ruta}")
    
    async def obtener_cookies(self):
        """Obtiene las cookies de la sesión actual"""
        cookies = await self.context.cookies()
        print(f"\n🍪 Cookies encontradas: {len(cookies)}")
        for cookie in cookies:
            print(f"  - {cookie['name']}: {cookie['value'][:20]}...")
        return cookies
    
    async def ejecutar_javascript(self, script: str):
        """Ejecuta JavaScript en la página"""
        print(f"⚡ Ejecutando JavaScript...")
        resultado = await self.page.evaluate(script)
        print(f"✅ Resultado: {resultado}")
        return resultado
    
    async def guardar_sesion(self, archivo: str = "state.json"):
        """Guarda la sesión actual en un archivo"""
        ruta = Path(archivo).resolve()
        storage = await self.context.storage_state()
        import json
        with open(ruta, 'w') as f:
            json.dump(storage, f, indent=2)
        print(f"💾 Sesión guardada en: {ruta}")
    
    async def cerrar(self):
        """Cierra el navegador"""
        try:
            if self.context:
                await self.context.close()
            if self.browser:
                await self.browser.close()
            if self.playwright:
                await self.playwright.stop()
            print("🔒 Navegador cerrado")
        except Exception as e:
            print(f"⚠️  Error al cerrar (ignorado): {e}")


from urllib.parse import unquote

@log_errors
async def esperar_carga_completa(page, timeout: int = 120):
    """
    Espera a que Outlook termine de cargar los contactos detectando el spinner.
    
    El spinner tiene dos estados:
    - Cargando: <div class="spinner spinnerAnimation">
    - Terminado: <div class="spinner"> (sin spinnerAnimation)
    
    Args:
        page: Página de Playwright
        timeout: Tiempo máximo de espera en segundos (default: 120)
    """
    spinner_selector = ".spinnerContainer .spinner"
    
    try:
        print("⏳ Esperando a que termine de cargar (detectando spinner)...")
        
        # Esperar a que el spinner aparezca Y tenga la clase spinnerAnimation (está cargando)
        # O esperar un poco si no aparece (máximo 5 segundos)
        try:
            await page.wait_for_selector(f"{spinner_selector}.spinnerAnimation", state="attached", timeout=5000)
            print("   🔄 Spinner detectado - cargando datos...")
        except:
            # Si no aparece el spinner en 5 segundos, probablemente ya cargó
            print("   ⚡ No se detectó spinner - los datos podrían estar listos")
            await asyncio.sleep(1)
            return
        
        # Ahora esperar a que DESAPAREZCA la clase spinnerAnimation (terminó de cargar)
        max_wait = timeout
        waited = 0
        check_interval = 0.5  # Revisar cada 0.5 segundos
        
        while waited < max_wait:
            # Verificar si el spinner todavía está animando
            spinner_animating = await page.locator(f"{spinner_selector}.spinnerAnimation").count()
            
            if spinner_animating == 0:
                print("   ✅ Carga completada (spinner desactivado)")
                # Espera adicional pequeña para asegurar que el DOM se actualizó
                await asyncio.sleep(1)
                return
            
            await asyncio.sleep(check_interval)
            waited += check_interval
            
            # Mostrar progreso cada 10 segundos
            if int(waited) % 10 == 0 and waited > 0:
                print(f"   ⏳ Todavía cargando... ({int(waited)}s)")
        
        print(f"   ⚠️ Timeout después de {timeout}s - continuando de todos modos")
        
    except Exception as e:
        print(f"   ⚠️ Error detectando spinner: {e} - usando espera de seguridad")
        await asyncio.sleep(3)

@log_errors
async def extraer_detalles_contacto(page, fila, nombre_contacto: str):
    """
    Hace clic en un contacto y extrae toda su información del panel de detalles.
    
    Args:
        page: Página de Playwright
        fila: Elemento de la fila del contacto
        nombre_contacto: Nombre del contacto para logging
    
    Returns:
        dict con toda la información del contacto
    """
    detalles = {
        'nombre': nombre_contacto,
        'email': '',
        'telefono_trabajo': '',
        'sip': '',
        'departamento': '',
        'compania': '',
        'oficina': '',
        'direccion': ''
    }
    
    try:
        # 1. Hacer clic en el contacto para abrir el panel de detalles
        print(f"   🖱️  Haciendo clic en: {nombre_contacto}")
        await fila.click()
        
        # 2. Esperar a que cargue el panel - buscar el campo "Compañía" como indicador
        try:
            # Esperamos hasta 20 segundos a que aparezca el campo Compañía
            await page.wait_for_selector('span._rpc_F1:has-text("Compañía:")', timeout=20000)
            print(f"   ✅ Panel de detalles cargado")
        except:
            # Si no aparece Compañía, esperamos un poco por si hay otros datos
            print(f"   ⚠️ Compañía no encontrada - esperando 3s adicionales")
            await asyncio.sleep(3)
        
        # 3. Obtener todo el texto del popup para extracción basada en texto
        popup_selector = "div._rpc_M, div[class*='_rpc']"  # Selector del panel de detalles
        popup = page.locator(popup_selector).first
        
        try:
            popup_text = await popup.inner_text()
        except:
            # Fallback: obtener de toda la página
            popup_text = await page.inner_text("body")
        
        lines = popup_text.split('\n')
        
        # 4. EXTRAER CAMPOS usando estrategia basada en labels (más robusta)
        
        # Email (personal, no el token)
        try:
            email_elements = await page.locator('span[title*="@"]').all()
            for elem in email_elements:
                try:
                    email_candidate = await elem.get_attribute('title')
                    if not email_candidate:
                        email_candidate = await elem.inner_text()
                    
                    email_candidate = email_candidate.strip()
                    
                    # Saltar tokens genéricos (AGM564@, ASP123@, etc.)
                    import re
                    if re.match(r'^(ASP|AGM|AEM|ADM)\d+@', email_candidate, re.I):
                        continue
                    
                    # Este parece ser un email personal
                    if '@' in email_candidate and '.' in email_candidate:
                        detalles['email'] = email_candidate
                        break
                except:
                    continue
        except Exception as e:
            print(f"   ⚠️ Error extrayendo email: {e}")
        
        # Teléfono (buscar "Trabajo:" y extraer la siguiente línea)
        try:
            for i, line in enumerate(lines):
                if 'Trabajo:' in line and 'Departamento' not in line:
                    # El teléfono está en la siguiente línea
                    if i + 1 < len(lines):
                        phone_value = lines[i + 1].strip()
                        # Limpiar solo dígitos y caracteres de teléfono
                        phone_clean = ''.join(c for c in phone_value if c.isdigit() or c in ['+', '-', ' ', '(', ')'])
                        if phone_clean.strip():
                            detalles['telefono_trabajo'] = phone_clean.strip()
                            break
        except Exception as e:
            print(f"   ⚠️ Error extrayendo teléfono: {e}")
        
        # SIP (buscar "MI:" y extraer la siguiente línea)
        try:
            for i, line in enumerate(lines):
                if line.strip() in ['MI:', 'MI', 'IM:', 'IM']:
                    # El SIP está en la siguiente línea
                    if i + 1 < len(lines):
                        sip_value = lines[i + 1].strip()
                        if sip_value.startswith('sip:'):
                            detalles['sip'] = sip_value
                            break
        except Exception as e:
            print(f"   ⚠️ Error extrayendo SIP: {e}")
        
        # Departamento
        try:
            for i, line in enumerate(lines):
                if 'Departamento:' in line:
                    # El valor puede estar en la misma línea o en la siguiente
                    value = line.replace('Departamento:', '').strip()
                    if not value and i + 1 < len(lines):
                        value = lines[i + 1].strip()
                    
                    if value and value.lower() not in ['directorio', 'directory', 'trabajo']:
                        detalles['departamento'] = value
                        break
        except Exception as e:
            print(f"   ⚠️ Error extrayendo departamento: {e}")
        
        # Compañía
        try:
            for i, line in enumerate(lines):
                if 'Compañía:' in line or 'Company:' in line:
                    # El valor puede estar en la misma línea o en la siguiente
                    value = line.replace('Compañía:', '').replace('Company:', '').strip()
                    if not value and i + 1 < len(lines):
                        value = lines[i + 1].strip()
                    
                    if value and value.lower() not in ['directorio', 'directory']:
                        detalles['compania'] = value
                        break
        except Exception as e:
            print(f"   ⚠️ Error extrayendo compañía: {e}")
        
        # Oficina
        try:
            for i, line in enumerate(lines):
                if 'Oficina:' in line or 'Office:' in line:
                    # El valor puede estar en la misma línea o en la siguiente
                    value = line.replace('Oficina:', '').replace('Office:', '').strip()
                    if not value and i + 1 < len(lines):
                        value = lines[i + 1].strip()
                    
                    if value and value.lower() not in ['directorio', 'directory']:
                        detalles['oficina'] = value
                        break
        except Exception as e:
            print(f"   ⚠️ Error extrayendo oficina: {e}")
        
        # Dirección (multilínea - siguientes 2-3 líneas después del label)
        try:
            for i, line in enumerate(lines):
                if 'Dirección profesional' in line or 'Business Address' in line:
                    # La dirección está en las siguientes 2-3 líneas
                    address_parts = []
                    for j in range(1, 4):  # Revisar siguientes 3 líneas
                        if i + j < len(lines):
                            line_text = lines[i + j].strip()
                            # Parar si encontramos otro label o línea vacía
                            if line_text and not any(x in line_text for x in ['Departamento', 'Compañía', 'Oficina', 'Trabajo:', 'MI:', 'Calendario']):
                                address_parts.append(line_text)
                            else:
                                break
                    
                    if address_parts:
                        detalles['direccion'] = ' '.join(address_parts)
                        break
        except Exception as e:
            print(f"   ⚠️ Error extrayendo dirección: {e}")
        
        print(f"   📋 Detalles: Email={detalles['email'][:30] if detalles['email'] else 'N/A'}, Dpto={detalles['departamento'][:30] if detalles['departamento'] else 'N/A'}, Cía={detalles['compania'][:30] if detalles['compania'] else 'N/A'}")
        
    except Exception as e:
        print(f"   ❌ Error extrayendo detalles: {e}")
    
    return detalles


@log_errors
async def scrape_outlook_contacts(page, max_contacts: int = 50):
    """Scrape de contactos de Outlook con scroll infinito y extracción de detalles
    
    Soporta reanudación: Si existe un Excel previo, continúa desde donde quedó.
    Usa metadata JSON para guardar el número de scrolls y reanudar más rápido.
    
    Args:
        page: Página de Playwright
        max_contacts: Número máximo TOTAL de contactos (incluyendo los ya guardados)
    """
    
    import pandas as pd
    from datetime import datetime
    import signal
    import sys
    import json
    
    # Selector de la fila
    row_selector = 'div[role="heading"]'
    
    # Buscar Excel más reciente en data/
    data_dir = Path("data")
    data_dir.mkdir(exist_ok=True)
    
    metadata_file = data_dir / "scraping_metadata.json"
    
    contactos_previos = []
    ultimo_nombre = None
    excel_path = None
    scroll_count_guardado = 0
    
    # Buscar el archivo Excel más reciente
    excel_files = list(data_dir.glob("contactos_organos_judiciales_*.xlsx"))
    if excel_files:
        excel_path = max(excel_files, key=lambda p: p.stat().st_mtime)
        print(f"📂 Archivo Excel encontrado: {excel_path}")
        
        # Cargar metadata si existe
        if metadata_file.exists():
            try:
                with open(metadata_file, 'r', encoding='utf-8') as f:
                    metadata = json.load(f)
                scroll_count_guardado = metadata.get('scroll_count', 0)
                print(f"📜 Metadata cargada - Scrolls previos: {scroll_count_guardado}")
            except Exception as e:
                print(f"⚠️ Error leyendo metadata: {e}")
        
        try:
            df_previo = pd.read_excel(excel_path)
            contactos_previos = df_previo.to_dict('records')
            print(f"✅ Cargados {len(contactos_previos)} contactos previos")
            
            if len(contactos_previos) > 0:
                ultimo_nombre = contactos_previos[-1]['nombre']
                print(f"🔄 Último contacto procesado: {ultimo_nombre}")
                print(f"📊 Faltan {max_contacts - len(contactos_previos)} contactos para llegar a {max_contacts}")
        except Exception as e:
            print(f"⚠️ Error leyendo Excel previo: {e}")
            contactos_previos = []
    else:
        print("📝 No se encontró Excel previo - comenzando desde cero")
    
    # Ajustar el límite según lo que ya tenemos
    contactos_faltantes = max(0, max_contacts - len(contactos_previos))
    if contactos_faltantes == 0:
        print(f"✅ Ya se alcanzó el límite de {max_contacts} contactos")
        return contactos_previos
    
    print(f"🎯 Extrayendo {contactos_faltantes} contactos adicionales...")
    
    # 1. ESPERA CRÍTICA: Esperar a que aparezca al menos un contacto
    print(f"⏳ Esperando a que cargue la lista del Directorio...")
    try:
        await page.wait_for_selector(row_selector, state="visible", timeout=30000)
        print("✅ Lista detectada")
    except Exception as e:
        print(f"❌ Error: La lista no cargó a tiempo. \nDetalle: {e}")
        await page.screenshot(path="debug_error_lista.png")
        return contactos_previos

    # 2. Si hay scrolls guardados, ir directamente a esa posición
    scroll_count_actual = 0
    
    if scroll_count_guardado > 0:
        print(f"\n⚡ Saltando directamente a scroll #{scroll_count_guardado}...")
        for i in range(scroll_count_guardado):
            await page.keyboard.press("PageDown")
            await asyncio.sleep(0.3)
            
            if i % 10 == 0 and i > 0:
                print(f"   Scrolleando... {i}/{scroll_count_guardado}")
        
        scroll_count_actual = scroll_count_guardado
        print(f"✅ Posicionado en scroll #{scroll_count_actual}")
        await asyncio.sleep(2)  # Esperar a que cargue
    
    # 3. Variables para el scraping
    contactos_nuevos = []
    contactos_procesados = set(c['nombre'] for c in contactos_previos)
    last_item_count = 0
    retries = 0
    max_retries = 15
    
    # 4. Handler para guardar automáticamente al interrumpir
    def guardar_y_salir(signum, frame):
        logger.warning("⚠️ Interrupción detectada - Guardando progreso...")
        print("\n\n⚠️ Interrupción detectada - Guardando progreso...")
        todos_contactos = contactos_previos + contactos_nuevos
        if todos_contactos:
            try:
                guardar_excel(todos_contactos, data_dir, excel_path)
                # Guardar metadata con scroll count
                guardar_metadata(metadata_file, contactos_previos + contactos_nuevos, scroll_count_actual)
                logger.info(f"💾 Progreso guardado - {len(todos_contactos)} contactos")
            except Exception as e:
                logger.error(f"❌ Error guardando al interrumpir: {e}")
                logger.error(traceback.format_exc())
        print("💾 Progreso guardado. Saliendo...")
        logger.info("🔚 Saliendo del scraper")
        sys.exit(0)
    
    signal.signal(signal.SIGINT, guardar_y_salir)
    signal.signal(signal.SIGTERM, guardar_y_salir)

    print(f"\n🎯 Iniciando extracción de {contactos_faltantes} contactos nuevos...")

    # 5. Loop principal de extracción
    while True:
        filas = await page.locator(row_selector).all()

        if not filas:
            print("⚠️ No se encontraron filas visibles.")
            break

        for fila in filas:
            try:
                raw_name = await fila.get_attribute("aria-label")
                nombre = raw_name.strip() if raw_name else "Desconocido"

                # Filtro: mayúsculas + coma
                if "," not in nombre or (not nombre.isupper()):
                    continue

                # Evitar duplicados
                if nombre in contactos_procesados:
                    continue
                
                contactos_procesados.add(nombre)
                
                total_actual = len(contactos_previos) + len(contactos_nuevos)
                print(f"\n🔍 Procesando ({total_actual + 1}/{max_contacts}): {nombre}")
                
                # Extraer detalles
                detalles = await extraer_detalles_contacto(page, fila, nombre)
                
                # FILTRO: Solo ORGANOS JUDICIALES
                if detalles['compania'] and 'ORGANOS JUDICIALES' in detalles['compania'].upper():
                    contactos_nuevos.append(detalles)
                    print(f"✅ Guardado ({total_actual + 1}/{max_contacts}) - ORGANOS JUDICIALES")
                else:
                    print(f"⏭️  Omitido - Compañía: {detalles['compania'] if detalles['compania'] else 'N/A'}")
                    contactos_procesados.remove(nombre)
                    continue
                
                # Verificar si alcanzamos el límite
                if len(contactos_previos) + len(contactos_nuevos) >= max_contacts:
                    print(f"\n🎯 ¡Límite alcanzado! Total: {max_contacts} contactos")
                    break

            except Exception as e:
                print(f"❌ Error procesando contacto: {e}")
        
        # Si alcanzamos el límite, salir
        if len(contactos_previos) + len(contactos_nuevos) >= max_contacts:
            break

        # Scroll
        try:
            if filas:
                await filas[-1].scroll_into_view_if_needed()
                await page.keyboard.press("PageDown")
                scroll_count_actual += 1  # Incrementar contador de scrolls
                await asyncio.sleep(0.5)
                await page.keyboard.press("PageDown")
                scroll_count_actual += 1  # Incrementar contador de scrolls
        except Exception as e:
            print(f"Error scroll: {e}")

        await esperar_carga_completa(page)

        # Verificación de fin
        current_count = len(contactos_nuevos)
        if current_count == last_item_count:
            retries += 1
            print(f"Sin nuevos contactos... Intento {retries}/{max_retries}")
            
            await page.keyboard.press("End")
            await asyncio.sleep(0.5)
            await page.keyboard.press("PageDown")
            scroll_count_actual += 1
            await esperar_carga_completa(page)
            
            if retries >= max_retries:
                print(f"⚠️ Terminado después de {retries} intentos.")
                break
        else:
            retries = 0
        
        last_item_count = current_count
    
    # 6. Guardar resultado final
    todos_contactos = contactos_previos + contactos_nuevos
    if todos_contactos:
        guardar_excel(todos_contactos, data_dir, excel_path)
        guardar_metadata(metadata_file, todos_contactos, scroll_count_actual)
    
    return todos_contactos

@log_errors
def guardar_metadata(metadata_file, contactos, scroll_count):
    """Guarda metadata de la sesión de scraping"""
    import json
    from datetime import datetime
    
    metadata = {
        'ultimo_contacto': contactos[-1]['nombre'] if contactos else None,
        'total_contactos': len(contactos),
        'scroll_count': scroll_count,
        'timestamp': datetime.now().isoformat()
    }
    
    logger.info(f"📝 Guardando metadata - Total: {len(contactos)}, Scrolls: {scroll_count}")
    with open(metadata_file, 'w', encoding='utf-8') as f:
        json.dump(metadata, f, indent=2, ensure_ascii=False)
    
    logger.info(f"✅ Metadata guardada en: {metadata_file}")
    print(f"📝 Metadata guardada - Scrolls: {scroll_count}")

@log_errors
def guardar_excel(contactos, data_dir, excel_path_existente=None):
    """Guarda contactos en Excel"""
    import pandas as pd
    from datetime import datetime
    
    if excel_path_existente and excel_path_existente.exists():
        # Actualizar el archivo existente
        excel_path = excel_path_existente
        logger.info(f"💾 Actualizando Excel existente: {excel_path}")
        print(f"\n💾 Actualizando Excel existente: {excel_path}")
    else:
        # Crear nuevo archivo con timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"contactos_organos_judiciales_{timestamp}.xlsx"
        excel_path = data_dir / excel_filename
        logger.info(f"📄 Creando nuevo Excel: {excel_path}")
        print(f"\n💾 Guardando {len(contactos)} contactos en nuevo Excel: {excel_path}")
    
    df = pd.DataFrame(contactos)
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Contactos')
        
        worksheet = writer.sheets['Contactos']
        for idx, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(col)
            ) + 2
            worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
    
    logger.info(f"✅ Excel guardado exitosamente: {excel_path} - {len(contactos)} contactos")
    print(f"✅ Excel guardado: {excel_path}")
    print(f"📊 Total contactos de ORGANOS JUDICIALES: {len(contactos)}")


@log_errors
async def main():
    """Función principal para debugging"""
    
    # Crear instancia del scraper
    # headless=False para ver el navegador mientras debugeas
    scraper = DebugScraper(headless=False)
    
    try:
        # 1. Iniciar sesión
        await scraper.iniciar_sesion("https://correoweb.madrid.org/owa/#path=/people")
        
        # Obtener la página después de inicializar
        page = scraper.page

        await page.wait_for_load_state("networkidle")

        directorio = page.get_by_label("Directorio", exact=True).locator("div").filter(has_text="Directorio").nth(1)
        await directorio.click()

        # Extraer contactos completos y guardar en Excel en data/
        await scrape_outlook_contacts(page, max_contacts=100)
        
        # await page.pause()

        await scraper.guardar_sesion()
    except KeyboardInterrupt:
        logger.warning("⚠️  Interrumpido por el usuario")
        print("\n\n⚠️  Interrumpido por el usuario")
    except Exception as e:
        logger.critical(f"❌ Error crítico en main: {e}")
        logger.critical(f"Traceback:\n{traceback.format_exc()}")
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        await scraper.cerrar()


if __name__ == "__main__":
    print("=" * 60)
    print("🐍 DEBUG SCRAPER - Listo para debugging")
    print("=" * 60)
    asyncio.run(main())
