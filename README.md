# Verificación de Correos OWA v2.0

Herramienta moderna de automatización Python que utiliza Playwright para interactuar con la interfaz webmail de Madrid (correoweb.madrid.org/owa) y extraer información de contactos de recipientes de correos electrónicos.

## 🚀 Características Principales

- **🔧 Arquitectura Moderna**: Estructura `src/` con paquetes Python estándar 2025
- **🧠 Automatización Inteligente**: Interactúa con OWA mediante Playwright
- **📊 Extracción Robusta**: Extrae información de contacto con múltiples métodos (DOM + regex)
- **⚡ Procesamiento por Lotes**: Procesamiento eficiente en lotes configurables
- **🖥️ Doble Interfaz**: CLI potente y GUI intuitiva
- **🔄 Sesión Persistente**: Reutilización de sesiones para evitar logins repetidos
- **📈 Registro Detallado**: Logging estructurado y seguimiento de progreso
- **🛠️ Configuración Flexible**: YAML + CLI + validación automática

## 📋 Requisitos

- Python 3.8+ (recomendado 3.11+)
- Navegador Chromium (instalado automáticamente por Playwright)

## ⚙️ Instalación y Configuración

### 1. Clonar el repositorio
```bash
git clone https://github.com/AndresGaibor/verificacion-correo.git
cd verificacion-correo
```

### 2. Crear entorno virtual
```bash
python3 -m venv .venv
source .venv/bin/activate  # En Windows: .venv\Scripts\activate
```

### 3. Instalar el paquete
```bash
# Modo desarrollo (recomendado)
pip install -e .

# O modo tradicional
pip install -r requirements.txt
playwright install chromium
```

### 4. Configurar archivo de configuración
```bash
# Si no existe, crear desde ejemplo
cp config.yaml.example config.yaml
```

Edita `config.yaml` con tu configuración específica.

### 5. Preparar archivo de datos
Crea un archivo Excel con los emails a procesar:
- Ubicación: `data/correos.xlsx` (o la ruta que prefieras)
- Columna A: Direcciones de correo (fila 2 en adelante)
- Fila 1: Encabezado "Correo"

### 6. Establecer sesión de autenticación
```bash
# Configurar sesión interactiva
verificacion-correo setup

# O alternativamente
python -m verificacion_correo setup
```

## 🔧 Uso

### Modo CLI (Línea de Comandos)

La nueva CLI incluye múltiples comandos y opciones:

```bash
# Procesar emails pendientes (comando por defecto)
verificacion-correo
verificacion-correo process

# Verificar estado y validación
verificacion-correo validate
verificacion-correo status

# Configurar sesión
verificacion-correo setup

# Usar archivo Excel específico
verificacion-correo --excel-file /path/to/emails.xlsx

# Cambiar tamaño de lote
verificacion-correo --batch-size 5

# Modo verbose y archivo de log
verificacion-correo -v --log-file processing.log

# Vista previa sin procesar
verificacion-correo --dry-run

# Forzar procesamiento incluso con sesión inválida
verificacion-correo --force
```

### Modo GUI (Interfaz Gráfica)

```bash
# Iniciar interfaz gráfica
verificacion-correo-gui
```

La GUI incluye:
- **📧 Pestaña de Procesamiento**: Control del procesamiento con barra de progreso
- **🔐 Pestaña de Sesión**: Gestión y estado de la sesión del navegador
- **⚙️ Pestaña de Configuración**: Información y acciones de configuración
- **📊 Registro en tiempo real**: Eventos y progreso detallados

## 📦 Estructura del Proyecto (v2.0)

```
verificacion-correo/
├── src/verificacion_correo/           # Paquete principal
│   ├── __init__.py                     # Inicialización del paquete
│   ├── __main__.py                     # Entry point CLI
│   ├── core/                          # Funcionalidades principales
│   │   ├── config.py                   # Gestión de configuración
│   │   ├── browser.py                  # Automatización del navegador
│   │   ├── extractor.py                # Extracción de contactos
│   │   ├── excel.py                    # Operaciones Excel
│   │   └── session.py                 # Gestión de sesiones
│   ├── cli/                           # Interfaz de línea de comandos
│   │   └── main.py                    # CLI principal
│   ├── gui/                           # Interfaz gráfica
│   │   └── main.py                    # GUI principal
│   └── utils/                         # Utilidades
│       └── logging.py                  # Configuración de logging
├── tests/                             # Suite de tests
│   ├── test_core/                     # Tests de funcionalidad core
│   └── test_integration/               # Tests de integración
├── config/                            # Archivos de configuración
│   └── default.yaml                   # Configuración por defecto
├── data/                              # Directorio de datos
│   └── correos.xlsx                   # Archivo Excel con emails
├── pyproject.toml                      # Configuración moderna del paquete
├── requirements.txt                    # Dependencias principales
├── requirements-dev.txt                # Dependencias de desarrollo
├── README.md                          # Documentación principal
└── CLAUDE.md                          # Guía para Claude Code
```

## 🔍 Comandos CLI Detallados

### `verificacion-correo process`
Procesa emails pendientes desde el archivo Excel configurado.

**Opciones:**
- `--excel-file PATH`: Archivo Excel específico
- `--batch-size N`: Tamaño de lote (default: 10)
- `--dry-run`: Vista previa sin ejecutar
- `--force`: Forzar procesamiento incluso con sesión inválida

### `verificacion-correo setup`
Configura sesión de navegador interactiva.
Abre una ventana para login manual y guarda el estado.

### `verificacion-correo validate`
Valida configuración y preparación del sistema:
- ✅ Archivo de configuración válido
- ✅ Sesión del navegador activa
- ✅ Archivo Excel accesible
- ✅ Emails pendientes disponibles

### `verificacion-correo status`
Muestra estado actual del sistema:
- 🌐 Información de la sesión (cookies, orígenes, validez)
- 📁 Estado del archivo Excel (existencia, tamaño, pendientes)
- ⚙️ Configuración actual

## 📊 Datos Extraídos

La herramienta extrae exitosamente **8 de 9 campos**:

- ✅ Email personal (ej: `nombre.apellido@madrid.org`)
- ✅ Teléfono de trabajo (ej: `916704092`)
- ✅ Dirección completa (ej: `C/ AYUNTAMIENTO, 5 28791 RIVAS-VACIAMADRID`)
- ✅ Departamento (ej: `OFICINA JUDICIAL MUNICIPAL`)
- ✅ Compañía (ej: `ORGANOS JUDICIALES`)
- ✅ Ubicación de oficina (ej: `RIVAS-VACIAMADRID`)
- ✅ Dirección SIP (ej: `sip:asp164@madrid.org`)
- ✅ Token de email (para identificación)
- ⚠️ **Nombre completo** (bloqueado por anti-scraping de Microsoft OWA)

### Limitación del Nombre de Usuario

Microsoft OWA implementa protección anti-bot que específicamente previene la extracción del nombre completo cuando se detecta automatización. Esta es una limitación conocida del lado del servidor y no puede evadirse con técnicas del lado del cliente.

## 🛠️ Desarrollo

### Configuración de Entorno de Desarrollo

```bash
# Instalar dependencias de desarrollo
pip install -r requirements-dev.txt

# Instalar en modo edición con dependencias de desarrollo
pip install -e ".[dev]"

# Formatear código
black src/ tests/

# Linting
ruff check src/ tests/

# Type checking
mypy src/

# Ejecutar tests
pytest

# Ejecutar tests con cobertura
pytest --cov=src --cov-report=html
```

### Scripts Útiles

```bash
# Validar configuración del paquete
python -m build --version

# Crear distribución
python -m build

# Publicar en PyPI (para maintainer)
python -m twine upload dist/*
```

## 🔄 Migración desde v1.0

Si vienes de la versión anterior:

1. **Reinstalar paquete**:
   ```bash
   pip install -e .
   ```

2. **Nuevos comandos**:
   - `python app.py` → `verificacion-correo`
   - `python copiar_sesion.py` → `verificacion-correo setup`
   - `python -m verificacion_correo.gui.main` → `verificacion-correo-gui`

3. **Configuración**:
   - `config.yaml` sigue siendo compatible
   - Nuevo `config/default.yaml` es la ubicación preferida

## 🏗️ Builds Automáticos (GitHub Actions)

El proyecto incluye workflows automáticos para:
- **CI/CD**: Tests en múltiples versiones de Python
- **Build Windows**: Ejecutable standalone para Windows
- **Release Automático**: Publicación automática en GitHub Releases

### Crear un Release

```bash
# Versión patch (v2.0.0 → v2.0.1)
./scripts/release.sh

# Versión minor (v2.0.0 → v2.1.0)
./scripts/release.sh minor

# Versión major (v2.0.0 → v3.0.0)
./scripts/release.sh major
```

## 📝 Licencia

Este proyecto está bajo licencia MIT. Ver el archivo `LICENSE` para más detalles.

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:

1. Fork del proyecto
2. Crear un feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit de los cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push al branch (`git push origin feature/AmazingFeature`)
5. Abrir un Pull Request

## 📞 Soporte

Si encuentras algún issue o tienes sugerencias:

- 🐛 Crea un issue en GitHub
- 📚 Revisa la documentación en `CLAUDE.md`
- 💬 Contacta al maintainer

---

**Desarrollado con ❤️ para automatización moderna de procesos de verificación de correos**

## 📈 Historial de Cambios (v2.0)

### ✨ Nuevas Características
- **Arquitectura Moderna**: Estructura `src/` con paquetes estándar
- **CLI Potente**: Múltiples comandos con validación y opciones
- **GUI Mejorada**: Interfaz más intuitiva y robusta
- **Logging Estructurado**: Configuración centralizada de logs
- **Tests**: Suite básica de tests unitarios e integración
- **Configuración Flexible**: Validación y fallbacks automáticos
- **Documentación**: README actualizado y guía de desarrollo

### 🔧 Mejoras Técnicas
- **Type Hints**: Anotaciones de tipo en todo el código
- **Error Handling**: Manejo robusto de errores
- **Performance**: Optimización del procesamiento por lotes
- **Maintenibility**: Código modular y bien documentado
- **Standards**: Cumplimiento con Python Packaging Authority 2025

### 🔄 Cambios Incompatibles
- Los comandos CLI han cambiado (ver sección de migración)
- La ubicación del código fuente ahora está en `src/`
- Se requiere Python 3.8+ (antes 3.11+)