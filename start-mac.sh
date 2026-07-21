#!/bin/bash
set -e

echo "========================================"
echo "  Verificación de Correo - Madrid"
echo "========================================"
echo ""

# Verificar que Python3 está disponible
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python3 no encontrado."
    echo "Instala Python desde: https://www.python.org/downloads/"
    echo "O usa Homebrew: brew install python@3.11"
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1)
echo "Python detectado: $PYTHON_VERSION"
echo ""

# Crear entorno virtual si no existe
if [ ! -d ".venv" ]; then
    echo "[1/3] Creando entorno virtual..."
    python3 -m venv .venv
    echo "      Entorno virtual creado."
    echo ""

    echo "[2/3] Instalando dependencias..."
    source .venv/bin/activate
    pip install -e . --quiet
    echo "      Dependencias instaladas."
    echo ""

    echo "[3/3] Instalando navegador Chromium para Playwright..."
    .venv/bin/playwright install chromium --quiet
    echo "      Navegador listo."
    echo ""
else
    source .venv/bin/activate
fi

echo "Iniciando interfaz gráfica..."
echo ""
python3 -m verificacion_correo.gui.main
