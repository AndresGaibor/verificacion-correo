# Detección de Autenticación Exitosa en Session Setup

**Fecha:** 2026-07-19
**Estado:** Aprobado

## Problema

`setup_interactive_session()` en `session.py` abre el navegador, espera 5 minutos fixed, y guarda el session state sin verificar si el login fue exitoso. Esto causa que:
- Sesiones fallidas se guarden igual
- El usuario no sepa si el login funcionó hasta que el proceso termine

## Solución

Modificar `setup_interactive_session()` para detectar dinámicamente cuándo el usuario completa la autenticación en ADFS (incluyendo MFA) y solo guardar la sesión si llega a `correoweb.madrid.org/owa/`.

## Arquitectura

### Flujo

1. Navegar a `page_url` (correoweb.madrid.org/owa/)
2. Detectar redirects → seguir hasta login ADFS si necesario
3. Loop de detección cada 2 segundos:
   - Si `correoweb.madrid.org/owa/` en URL → auth exitosa → guardar sesión → return True
   - Si `login` o `signin` en URL →仍在 formulario → seguir esperando
   - Timeout 180s → auth fallida → no guardar → return False
4. El usuario puede cerrar browser manualmente → loop detecta page cerrado → sale

### Detalles de Implementación

- **Módulo:** `src/verificacion_correo/core/session.py`
- **Función:** `SessionManager.setup_interactive_session()`
- **Timeout:** 180 segundos (3 minutos)
- **Check interval:** 2 segundos
- **URL éxito:** `correoweb.madrid.org/owa/` presente en page.url

### Cambios de Código

```python
import time

def setup_interactive_session(self) -> bool:
    # ... existing browser launch code ...

    page = context.new_page()
    page.goto(self.config.page_url)

    # Detection loop for successful authentication
    start_time = time.time()
    TIMEOUT = 180  # 3 minutes
    CHECK_INTERVAL = 2  # seconds

    logger.info("Waiting for authentication... (timeout: 3 minutes)")

    while True:
        # Check if browser was closed by user
        if page.is_closed():
            logger.info("Browser closed by user")
            break

        current_url = page.url.lower()

        # Success: OWA page reached
        if 'correoweb.madrid.org/owa/' in current_url:
            logger.info("Authentication successful!")
            self._ensure_session_directory()
            context.storage_state(path=str(self.session_file))
            context.close()
            browser.close()
            return True

        # Still on login page - keep waiting
        time.sleep(CHECK_INTERVAL)

        # Check timeout
        if time.time() - start_time >= TIMEOUT:
            logger.warning("Authentication timeout (3 minutes)")
            break

    # Cleanup on failure
    try:
        context.close()
        browser.close()
    except Exception:
        pass

    return False
```

## Comportamiento

| Escenario | Resultado |
|----------|-----------|
| Login exitoso + MFA | Browser cierra, sesión guardada, return True |
| Login fallido (bad credentials) | Timeout 3 min, sesión NO guardada, return False |
| Usuario cierra browser manual | Browser cierra, sesión NO guardada, return False |
| Page crash/error | Browser cierra, sesión NO guardada, return False |

## Auto-login con Credenciales (Extensión)

### Problema Adicional
El usuario quiere que si tiene credenciales configuradas, el sistema llene automáticamente el formulario de login ADFS y haga clic en "Iniciar sesión", ahorrando tiempo手动.

### Solución

Agregar configuración de credenciales opcionales. Si están presentes, el sistema:
1. Navega a la página de login ADFS
2. Auto-llena email y password del formulario
3. Hace clic en "Iniciar sesión"
4. El usuario introduce el código MFA manualmente
5. Loop de detección espera a que llegue a OWA
6. Guarda sesión

### Fuentes de Credenciales (en orden de prioridad)

1. **Variables de entorno**: `ADFS_USERNAME`, `ADFS_PASSWORD`
2. **YAML config** bajo `auth: { username, password }`

### Configuración

```yaml
# config/default.yaml
auth:
  username: ""  # O env var ADFS_USERNAME
  password: ""  # O env var ADFS_PASSWORD
```

```bash
# .env (o exportados)
export ADFS_USERNAME="usuario@madrid.org"
export ADFS_PASSWORD="miPassword"
```

### Selectores ADFS Login Form

```python
LOGIN_FORM = {
    'username_input': '#userNameInput',
    'password_input': '#passwordInput',
    'submit_btn': '#submitButton',
    'form': '#loginForm',
    'error_display': '#errorText'
}
```

### Flujo Modificado

1. Navegar a `page_url`
2. Si hay redirect a ADFS login (`signin` o `login` en URL):
   - Si credenciales disponibles → auto-llenar y submit
   - Si no hay credenciales → esperar login manual
3. Loop de detección cada 2s:
   - Si `correoweb.madrid.org/owa/` en URL → guardar sesión → return True
   - Si sigue en login sin credenciales → esperar usuario manual
   - Timeout 180s → no guardar → return False

### Cambios en Config

Nuevo dataclass `AuthConfig`:
```python
@dataclass
class AuthConfig:
    username: Optional[str] = None
    password: Optional[str] = None

    def has_credentials(self) -> bool:
        return bool(self.username and self.password)
```

Y en `Config._initialize_components()`:
```python
# Auth credentials - from env vars first, then YAML
env_username = os.environ.get('ADFS_USERNAME', '')
env_password = os.environ.get('ADFS_PASSWORD', '')
yaml_username = self._config_data.get('auth', {}).get('username', '')
yaml_password = self._config_data.get('auth', {}).get('password', '')

self.auth = AuthConfig(
    username=env_username or yaml_username or None,
    password=env_password or yaml_password or None
)
```

## Testing

- Verificar que al autenticar correctamente, la sesión se guarda
- Verificar que si no se autentica, no se guarda sesión
- Verificar que el timeout funciona a los 3 minutos
- Verificar auto-login con credenciales configuradas
- Verificar fallback a login manual cuando no hay credenciales
