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

__version__ = "2.0.0"
__author__ = "Andres Gaibor"
__email__ = "andres@example.com"

# Lazy imports para evitar cargar tkinter/playwright al importar el paquete
# Los módulos se importan bajo demanda desde sus respectivos entry points

__all__ = [
    "config",
    "browser",
    "extractor",
    "excel",
    "session",
    "cli_main",
    "gui_main",
]