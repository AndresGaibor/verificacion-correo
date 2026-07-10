@echo off
REM ============================================================
REM  setup_windows.bat - Setup automatico para Windows
REM  Crea venv, instala dependencias, configura entorno
REM ============================================================

setlocal enabledelayedexpansion

echo.
echo ============================================================
echo  Verificacion de Correos OWA - Setup para Windows
echo ============================================================
echo.

REM 1. Verificar Python
echo [1/5] Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [X] Python no esta instalado o no esta en PATH
    echo     Descarga Python 3.10+ desde https://www.python.org/downloads/
    echo     Durante la instalacion marca "Add Python to PATH"
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version') do echo     Detectado: %%v
echo.

REM 2. Crear entorno virtual
echo [2/5] Creando entorno virtual .venv...
if exist .venv (
    echo     .venv ya existe, saltando creacion
) else (
    python -m venv .venv
    if errorlevel 1 (
        echo [X] Error creando venv
        pause
        exit /b 1
    )
    echo     .venv creado OK
)
echo.

REM Activar venv para los siguientes pasos
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo [X] No se pudo activar .venv
    pause
    exit /b 1
)

REM 3. Actualizar pip e instalar dependencias
echo [3/5] Instalando dependencias Python...
python -m pip install --upgrade pip --quiet
python -m pip install playwright openpyxl pyyaml --quiet
if errorlevel 1 (
    echo [X] Error instalando dependencias
    pause
    exit /b 1
)
echo     playwright, openpyxl, pyyaml instalados OK
echo.

REM 4. Instalar Chromium para Playwright
echo [4/5] Instalando navegador Chromium...
python -m playwright install chromium
if errorlevel 1 (
    echo [!] Warning: No se pudo instalar Chromium automaticamente
    echo     Intenta manualmente: python -m playwright install chromium
)
echo.

REM 5. Crear config.yaml desde ejemplo
echo [5/5] Configurando archivos base...
if not exist config.yaml (
    if exist config.yaml.example (
        copy /Y config.yaml.example config.yaml >nul
        echo     config.yaml creado desde ejemplo
        echo     IMPORTANTE: Edita config.yaml con la URL real de OWA
    ) else (
        echo [X] No se encuentra config.yaml.example
    )
) else (
    echo     config.yaml ya existe, no se sobrescribe
)

REM Crear data/ si no existe
if not exist data mkdir data
echo     Directorio data/ listo

REM Crear plantilla Excel si no existe data/correos.xlsx
if not exist data\correos.xlsx (
    echo     Creando plantilla Excel data\correos.xlsx...
    python -c "import openpyxl; wb=openpyxl.Workbook(); ws=wb.active; ws.title='Contactos'; ws.append(['Correo','Status','Nombre','Email Personal','Telefono','SIP','Direccion','Departamento','Compania','Oficina']); wb.save('data/correos.xlsx'); print('Plantilla creada')"
)

echo.
echo ============================================================
echo  Setup completado!
echo ============================================================
echo.
echo  Proximos pasos:
echo  1. Edita config.yaml con la URL de OWA
echo  2. Edita data\correos.xlsx y anade correos
echo  3. Doble clic en start.bat para iniciar la GUI
echo  4. Desde la GUI, ve a la pestana "Sesion" para hacer login manual
echo.
echo  Alternativamente desde CMD:
echo     .venv\Scripts\activate
echo     python -m verificacion_correo.gui.main    REM GUI completa
echo     verificacion-correo help                  REM CLI
echo.
pause
