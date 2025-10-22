# Verificación de Correos OWA

Herramienta de automatización Python que utiliza Playwright para interactuar con la interfaz webmail de Madrid (correoweb.madrid.org/owa) y extraer información de contactos de recipientes de correos electrónicos.

## 🚀 Características

- **Automatización Inteligente**: Interactúa con OWA (Outlook Web Access) mediante Playwright
- **Extracción de Datos**: Extrae automáticamente información de contacto de emails
- **Procesamiento por Lotes**: Procesa múltiples correos en lotes para mejor rendimiento
- **Interfaz GUI**: Incluye interfaz gráfica para facilitar el uso
- **Sesión Persistente**: Reutiliza sesiones de autenticación para evitar logins repetidos
- **Exportación Excel**: Guarda resultados en archivos Excel con formato estructurado

## 📋 Requisitos

- Python 3.11+
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

### 3. Instalar dependencias
```bash
pip install -r requirements.txt
playwright install chromium
```

### 4. Configurar el archivo de configuración
```bash
# Copiar el archivo de ejemplo y configurarlo
cp config.yaml.example config.yaml
```

Edita `config.yaml` con tu configuración:
- URL del servidor OWA
- Emails por defecto
- Selectores CSS para la interfaz
- Tiempos de espera

### 5. Crear archivo de datos
Crea un archivo Excel `data/correos.xlsx` con los emails a procesar:
- Columna A: Direcciones de correo (fila 2 en adelante)
- Fila 1: Encabezado "Correo"

### 6. Establecer sesión de autenticación
```bash
python copiar_sesion.py
# Sigue las instrucciones para login manual y presiona ENTER para guardar sesión
```

## 🔧 Uso

### Modo CLI (Línea de Comandos)
```bash
# Procesar todos los emails del Excel
python app.py
```

### Modo GUI (Interfaz Gráfica)
```bash
# Iniciar interfaz gráfica
python iniciar_gui.py
```

## 📦 Builds Automáticos (GitHub Actions)

El proyecto incluye un workflow de GitHub Actions que automáticamente:

- Compila un ejecutable **Windows standalone** cuando se crea un tag `v*`
- Cachea dependencias para builds rápidos
- Publica releases automáticos en GitHub
- Genera release notes automáticamente

### Crear un Release

Usa el script de release automatizado:

```bash
# Incrementar versión patch (v1.0.0 → v1.0.1)
./scripts/release.sh

# Incrementar versión minor (v1.0.0 → v1.1.0)
./scripts/release.sh minor

# Incrementar versión major (v1.0.0 → v2.0.0)
./scripts/release.sh major

# Ver comandos sin ejecutar (dry-run)
./scripts/release.sh --dry-run

# Forzar sin confirmación
./scripts/release.sh minor --force

# Release con mensaje personalizado
./scripts/release.sh patch -m "Hotfix crítico"
```

El script automáticamente:
1. Calcula la nueva versión semántica
2. Crea el tag localmente
3. Sube el tag a GitHub
4. Dispara el workflow de build

### Opciones del Script de Release

```bash
./scripts/release.sh [TIPO] [OPCIONES]

Tipos:
  patch    Incrementa versión patch (por defecto)
  minor    Incrementa versión minor
  major    Incrementa versión major

Opciones:
  --force         No pedir confirmación
  --dry-run       Mostrar comandos sin ejecutar
  --allow-dirty   Permitir cambios sin commitear
  -m, --message   Mensaje personalizado para el tag
  -h, --help      Mostrar ayuda
```

## 🏗️ Estructura del Proyecto

```
verificacion-correo/
├── .github/workflows/
│   └── build-windows.yml          # Workflow GitHub Actions
├── scripts/
│   └── release.sh                 # Script de release automatizado
├── app.py                         # Script principal (CLI)
├── iniciar_gui.py                 # Iniciador de interfaz gráfica
├── config.py                      # Gestión de configuración
├── excel_reader.py                # Lectura de datos desde Excel
├── excel_writer.py                # Escritura de resultados a Excel
├── browser_automation.py         # Automatización del navegador
├── contact_extractor.py           # Extracción de información de contacto
├── copiar_sesion.py               # Gestión de sesiones
├── gui.py                         # Interfaz gráfica principal
├── gui_config_manager.py          # Gestión de configuración en GUI
├── gui_runner.py                  # Ejecución de procesos en GUI
├── gui_session_manager.py         # Gestión de sesiones en GUI
├── config.yaml.example            # Plantilla de configuración
├── requirements.txt               # Dependencias Python
├── data/                          # Directorio de datos
│   └── correos.xlsx              # Archivo Excel con emails
└── state.json                     # Archivo de sesión (ignorado por git)
```

## 🔍 Información Extraída

La herramienta extrae la siguiente información de cada contacto:

- ✅ Email personal
- ✅ Teléfono de trabajo
- ✅ Dirección postal completa
- ✅ Departamento
- ✅ Compañía
- ✅ Ubicación de oficina
- ✅ Dirección SIP
- ✅ Token email
- ⚠️ Nombre completo (limitado por anti-scraping de Microsoft OWA)

## ⚠️ Limitaciones Conocidas

Microsoft OWA implementa medidas anti-scraping que **previenen la extracción del nombre completo** cuando se detecta automatización. Todos los demás campos se extraen correctamente.

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

- Crea un issue en GitHub
- Revisa la documentación en `CLAUDE.md`
- Contacta al maintainer

---

**Desarrollado con ❤️ para automatización de procesos de verificación de correos**