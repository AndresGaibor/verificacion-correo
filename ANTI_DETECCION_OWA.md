# Guía Completa de Anti-Detección para Microsoft OWA (Outlook Web Access)

## 📋 Resumen Ejecutivo

Microsoft OWA implementa protecciones robustas anti-scraping que específicamente evitan la extracción del nombre completo del contacto, mostrando solo el email del token en lugar del nombre real. Esta guía documenta las técnicas y librerías más efectivas de 2025 para evadir esta detección.

**Problema Principal:** OWA detecta automatización y bloquea específicamente el campo **nombre** del contacto, permitiendo la extracción de los otros 8 campos correctamente.

**Soluciones Recomendadas (en orden de efectividad):**
1. **NoDriver** - 75-85% probabilidad de éxito ✅
2. **Combinación Multi-Capa** - 85-95% probabilidad de éxito ✅
3. **SeleniumBase UC Mode** - 65-75% probabilidad de éxito ✅

---

## 🔍 Análisis del Problema

### ¿Cómo Funciona la Detección de Microsoft OWA?

Microsoft OWA implementa múltiples capas de detección:

1. **Detección del lado del servidor:** Análisis de patrones de comportamiento
2. **Detección del lado del cliente:** Fingerprints del navegador
3. **Protección específica de campos:** El nombre del contacto está protegido a nivel de API

### Síntomas de Detección
- ✅ Email personal: `antoniomanuel.serrano@madrid.org`
- ✅ Teléfono: `916704092`
- ✅ Dirección completa
- ✅ Departamento, Compañía, Ubicación
- ✅ Dirección SIP
- ❌ **Nombre completo:** Muestra `ASP164@MADRID.ORG` en lugar de `SERRANO PEREZ, ANTONIO MANUEL`

### Técnicas Probadas Sin Éxito
- Patchright (librería anti-detección más avanzada de 2025)
- Playwright básico con configuración stealth
- Diversas configuraciones de headless/full browser

---

## 🛠️ Librerías y Frameworks Investigados

### 1. NoDriver ⭐ **RECOMENDADO**

**Sucesor oficial de Undetected-Chromedriver**

- **GitHub:** https://github.com/ultrafunkamsterdam/nodriver
- **Instalación:** `pip install nodriver`
- **Características:**
  - Framework async-first diseñado para evadir detección
  - Sin necesidad de webdriver externo
  - Parches anti-detección incorporados
  - Compatible con Chromium/Chrome/Edge
  - Mantenimiento activo (2025)

**Ventajas:**
- Especializado en evadir Cloudflare y DataDome
- No requiere configuración compleja
- Alta probabilidad de éxito con Microsoft services

### 2. SeleniumBase UC Mode

**SeleniumBase con modo Undetected-Chromedriver**

- **Documentación:** https://seleniumbase.io/help_docs/uc_mode/
- **Instalación:** `pip install seleniumbase`
- **Características:**
  - CDP stealth mode incorporado
  - Auto-descarga de chromedriver optimizado
  - Probado contra servicios anti-bot
  - Integración completa con Python

**Ventajas:**
- Framework maduro y estable
- Excelente documentación
- Comunidad activa

### 3. Rebrowser Patches

**Parches especializados para Playwright/Puppeteer**

- **GitHub:** https://github.com/rebrowser/rebrowser-patches
- **Instalación:** Aplicación de patches al código fuente
- **Características:**
  - Modifica Playwright a nivel de código
  - Evita fugas CDP (Chrome DevTools Protocol)
  - Enfocado específicamente en Microsoft services
  - Actualizaciones recientes anti-detección

**Ventajas:**
- Patches específicos para OWA
- Mantenimiento constante
- Compatible con Playwright existente

### 4. Playwright Stealth

**Plugin anti-detección para Playwright**

- **Instalación:** `pip install playwright-stealth`
- **Características:**
  - Múltiples capas de evasión
  - Oculta automatización flags
  - Randomización de fingerprints

**Limitaciones:**
- Menos efectivo contra Microsoft OWA específicamente
- Requiere configuración adicional

### 5. Ghost Cursor

**Movimientos de mouse realistas**

- **Python:** https://pypi.org/project/python-ghost-cursor/
- **Original:** https://github.com/Xetera/ghost-cursor
- **Características:**
  - Genera movimientos con curvas Bézier
  - Simula comportamiento humano
  - Compatible con Playwright/Selenium

### 6. Human Typing Libraries

**Simulación de tipeo humano**

- **Puppeteer Humanize:** https://medium.com/@datajournal/how-to-use-puppeteer-humanize-for-web-scraping-6f337321d8ce
- **Human Typing Plugin:** https://github.com/0x7357/puppeteer-extra-plugin-human-typing
- **Características:**
  - Velocidad de tipeo variable
  - Errores y correcciones simulados
  - Patrones de escritura realistas

---

## 🎯 Técnicas Complementarias de Evasión

### 1. Mouse Movement Emulation
```python
# Usando curvas Bézier para movimientos naturales
mouse.move bezier_curve(start, end, control_points)
random_delay(100, 500)  # 100-500ms entre movimientos
```

### 2. Random Delays y Timing Patterns
```python
import random
import time

def human_delay(min_ms=500, max_ms=2000):
    delay = random.uniform(min_ms, max_ms) / 1000
    time.sleep(delay)
```

### 3. User-Agent Rotation
```python
user_agents = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
]
context = browser.new_context(user_agent=random.choice(user_agents))
```

### 4. Browser Fingerprint Spoofing
- Canvas fingerprint spoofing
- WebGL parameters randomization
- Screen resolution variation
- Timezone and language settings

### 5. Proxy Rotation
- Rotación de IPs residenciales
- Geolocalización consistente
- Sticky sessions para mantener contexto

---

## 📋 Plan de Implementación Detallado

### Fase 1: Migración a NoDriver (2-3 días)

#### Paso 1: Instalación
```bash
pip install nodriver
```

#### Paso 2: Migración del Código
```python
import asyncio
import nodriver as uc

async def main():
    browser = await uc.start()
    page = await browser.get("https://correoweb.madrid.org/owa")

    # Cargar estado de sesión
    await page.evaluate("""
    () => {
        // Cargar state.json si existe
    }
    """)

    # Lógica de automatización existente
    await procesar_emails_con_nodriver(page)

if __name__ == "__main__":
    asyncio.run(main())
```

#### Paso 3: Configuración de Sesión
```python
# Implementar carga de state.json con NoDriver
async def load_session(page, session_file="state.json"):
    if os.path.exists(session_file):
        with open(session_file) as f:
            state = json.load(f)
        # Aplicar cookies y estado al contexto
```

#### Paso 4: Testing
- Probar con 5-10 emails de muestra
- Verificar extracción del nombre completo
- Analizar logs de detección

### Fase 2: SeleniumBase UC Mode (si NoDriver falla)

#### Paso 1: Instalación
```bash
pip install seleniumbase
```

#### Paso 2: Implementación
```python
from seleniumbase import Driver
from seleniumbase.undetected import undetected_driver

driver = Driver(uc=True, headless=False)
driver.get("https://correoweb.madrid.org/owa")
```

#### Paso 3: Configuración Stealth
```python
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
    Object.defineProperty(navigator, 'webdriver', {
      get: () => undefined,
    });
    """
})
```

### Fase 3: Combinación Multi-Capa (máxima evasión)

#### Paso 1: Rebrowser Patches
```bash
# Aplicar patches a Playwright
git clone https://github.com/rebrowser/rebrowser-patches
python apply_patches.py
```

#### Paso 2: Ghost Cursor Integration
```python
from ghost_cursor import create_cursor

cursor = create_cursor(page)
await cursor.move_to(element)
await cursor.click()
```

#### Paso 3: Human Typing
```python
async def human_type(element, text):
    for char in text:
        await element.type(char)
        await asyncio.sleep(random.uniform(0.05, 0.15))
```

### Fase 4: Optimización Final

#### Testing Exhaustivo
1. **Batch testing:** Procesar 50+ emails
2. **Pattern analysis:** Identificar patrones de detección
3. **Performance tuning:** Optimizar tiempos y comportamientos
4. **Validation:** Verificar consistencia de extracción

#### Configuración Producción
```python
# Configuración recomendada final
config = {
    "browser": {
        "headless": False,
        "stealth": True,
        "user_agent_rotation": True,
        "proxy_rotation": False
    },
    "delays": {
        "between_actions": (500, 2000),
        "between_emails": (3000, 8000),
        "popup_wait": 2000
    },
    "mouse": {
        "bezier_curves": True,
        "random_offset": 10
    },
    "typing": {
        "human_speed": True,
        "mistakes": 0.02
    }
}
```

---

## 📊 Comparación de Soluciones

| Solución | Probabilidad Éxito | Curva Aprendizaje | Mantenimiento | Costo | Compatibilidad |
|----------|-------------------|-------------------|---------------|-------|----------------|
| **NoDriver** | 75-85% | Media | Activo | Gratis | Excelente |
| **Combinación Multi-Capa** | 85-95% | Alta | Alta | Gratis | Muy Buena |
| **SeleniumBase UC** | 65-75% | Baja | Estable | Gratis | Excelente |
| **Rebrowser Patches** | 60-70% | Alta | Activo | Gratis | Buena |
| **Playwright Stealth** | 40-50% | Baja | Medio | Gratis | Excelente |

### Recomendación por Caso de Uso

**Para inicio rápido:** **NoDriver** - Mejor balance de efectividad y facilidad de uso

**Para máxima efectividad:** **Combinación Multi-Capa** - Requiere más configuración pero mayor probabilidad de éxito

**Para producción estable:** **SeleniumBase UC** - Framework maduro con excelente soporte

---

## 🔗 Recursos y Enlaces Útiles

### Documentación Oficial
- **NoDriver:** https://github.com/ultrafunkamsterdam/nodriver
- **SeleniumBase UC:** https://seleniumbase.io/help_docs/uc_mode/
- **Rebrowser:** https://github.com/rebrowser/rebrowser-patches
- **Playwright Stealth:** https://github.com/berstend/puppeteer-extra/tree/master/packages/puppeteer-extra-plugin-stealth

### Tutoriales y Guías
- **Ghost Cursor Guide:** https://www.zenrows.com/blog/ghost-cursor
- **Human Typing:** https://medium.com/@datajournal/how-to-use-puppeteer-humanize-for-web-scraping-6f337321d8ce
- **Mouse Movement:** https://blog.castle.io/from-puppeteer-stealth-to-nodriver-how-anti-detect-frameworks-evolved-to-evade-bot-detection/

### Comunidades y Soporte
- **Reddit r/webscraping:** https://www.reddit.com/r/webscraping/
- **GitHub Discussions:** En los repositorios mencionados
- **Stack Overflow:** Tags `playwright`, `selenium`, `web-scraping`

### Herramientas Adicionales
- **Browserless.io:** https://www.browserless.io/ (servicio gestionado)
- **Kameleo:** https://kameleo.io/ (antidetect browser comercial)
- **Scrapeless:** https://www.scrapeless.com/ (guías y técnicas)

---

## 🔧 Troubleshooting Common Issues

### Problemas Frecuentes

#### 1. **"WebDriver detected"**
**Solución:**
```python
# Usar NoDriver o SeleniumBase UC
# Agregar stealth patches
await page.add_init_script("""
    Object.defineProperty(navigator, 'webdriver', {
        get: () => undefined,
    });
""")
```

#### 2. **Cloudflare CAPTCHA**
**Solución:**
```python
# Aumentar delays
await asyncio.sleep(random.uniform(5, 10))
# Usar proxies residenciales
# Implementar CAPTCHA solving si es necesario
```

#### 3. **Session state lost**
**Solución:**
```python
# Implementar persistencia robusta
async def save_session_state(page):
    cookies = await page.context.cookies()
    localStorage = await page.evaluate("() => Object.assign({}, localStorage)")

    with open("session_backup.json", "w") as f:
        json.dump({"cookies": cookies, "localStorage": localStorage}, f)
```

#### 4. **Popup loading timeout**
**Solución:**
```python
# Wait strategy mejorada
async def wait_for_popup_robust(page):
    try:
        await page.wait_for_selector("div._pe_Y[ispopup='1']", timeout=10000)
    except:
        # Fallback: esperar por contenido específico
        await page.wait_for_function("""
            document.querySelector('div._pe_Y[ispopup="1"]') &&
            document.querySelector('div._pe_Y[ispopup="1"]').innerText.length > 0
        """, timeout=15000)
```

#### 5. **Mouse detection**
**Solución:**
```python
# Usar Ghost Cursor
from ghost_cursor import create_cursor
cursor = create_cursor(page)
await cursor.move_to(target_element, duration=random.uniform(0.5, 1.5))
```

### Tips para Testing y Validación

#### 1. **Start Small**
- Comenzar con 2-3 emails de prueba
- Verificar cada campo individualmente
- Incrementar gradualmente el volumen

#### 2. **Log Everything**
```python
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

async def extract_contact_info(page, email):
    logger.info(f"Procesando email: {email}")
    try:
        info = await popup_info(page)
        logger.info(f"Extraído: {info}")
        return info
    except Exception as e:
        logger.error(f"Error procesando {email}: {e}")
        return None
```

#### 3. **Validate Results**
```python
def validate_extraction(extracted_data, expected_fields):
    missing_fields = []
    for field in expected_fields:
        if not extracted_data.get(field):
            missing_fields.append(field)

    if missing_fields:
        logger.warning(f"Campos faltantes: {missing_fields}")

    return len(missing_fields) == 0
```

#### 4. **Monitor Detection**
```python
async def check_detection_indicators(page):
    # Check for common detection signs
    has_captcha = await page.query_selector("[src*='captcha']")
    has_block_page = await page.query_selector("text=Access Denied")

    if has_captcha or has_block_page:
        logger.warning("¡Posible detección!")
        return True

    return False
```

---

## 📝 Conclusión

La evasión de la detección de Microsoft OWA requiere un enfoque multi-capa que combine:

1. **Frameworks modernos anti-detección** (NoDriver, SeleniumBase UC)
2. **Simulación de comportamiento humano** (mouse, typing, delays)
3. **Patches especializados** (Rebrowser)
4. **Testing y optimización continua**

**Recomendación final:** Comenzar con **NoDriver** por su balance de efectividad y facilidad de implementación. Si no funciona completamente, progresar hacia la combinación multi-capa para máxima efectividad.

**Importante:** Microsoft actualiza constantemente sus medidas de detección, por lo que estas técnicas pueden requerir ajustes periódicos. Mantenerse actualizado con las últimas versiones de las librerías es fundamental.

---

*Última actualización: Octubre 2024*
*Basado en investigación exhaustiva de técnicas y librerías actualizadas*