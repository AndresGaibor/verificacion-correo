"""
Verificación de Correo - Herramienta de automatización para extracción de contactos OWA

Este paquete proporciona funcionalidades para automatizar la extracción de información
de contacto desde la interfaz web de correo de Madrid (OWA) usando Playwright.

Main components:
- core: Funcionalidades principales de automatización y extracción
- cli: Interfaz de línea de comandos
- gui: Interfaz gráfica de usuario
- utils: Utilidades comunes
"""

__version__ = "1.0.0"
__author__ = "Andres Gaibor"
__email__ = "andres@example.com"

from .core import config, browser, extractor, excel, session
from .cli import main as cli_main
from .gui import main as gui_main

__all__ = [
    "config",
    "browser",
    "extractor",
    "excel",
    "session",
    "cli_main",
    "gui_main",
]