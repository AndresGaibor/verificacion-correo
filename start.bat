@echo off
title Verificacion de Correo - Madrid
echo ========================================
echo   Verificacion de Correo - Madrid
echo ========================================
echo.

REM Verificar que Python esta disponible
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no encontrado.
    echo Descarga Python desde: https://www.python.org/downloads/
    echo Marca "Add Python to PATH" durante la instalacion.
    echo.
    pause
    exit /b 1
)

REM Crear entorno virtual si no existe
if not exist ".venv\Scripts\activate" (
    echo [1/3] Creando entorno virtual...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [ERROR] No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
    echo       Entorno virtual creado.
    echo.

    echo [2/3] Instalando dependencias...
    call .venv\Scripts\activate
    pip install -e . >nul 2>&1
    if %errorlevel% neq 0 (
        echo [ERROR] No se pudieron instalar las dependencias.
        pause
        exit /b 1
    )
    echo       Dependencias instaladas.
    echo.

    echo [3/3] Instalando navegador Chromium para Playwright...
    .venv\Scripts\playwright install chromium >nul 2>&1
    echo       Navegador listo.
    echo.
) else (
    call .venv\Scripts\activate
)

echo Iniciando interfaz grafica...
echo.
python -m verificacion_correo.gui.main
pause
