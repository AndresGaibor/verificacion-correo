#!/usr/bin/env python3
"""
Script de lanzamiento para la interfaz gráfica.
Facilita el inicio de la GUI con verificación de prerrequisitos.
"""

import sys
import os
from tkinter import messagebox
import tkinter as tk


def verificar_prerrequisitos():
    """Verifica que los prerrequisitos básicos estén cumplidos."""
    errores = []

    # Verificar archivo de sesión
    if not os.path.exists("state.json"):
        errores.append(
            "❌ No se encuentra el archivo de sesión (state.json)\n"
            "   Ejecuta primero: python scraper_gui.py o usa la pestaña 'Sesión' en esta GUI"
        )

    # Verificar archivo de configuración
    if not os.path.exists("config.yaml"):
        if os.path.exists("config.yaml.example"):
            errores.append(
                "⚠️ No se encuentra config.yaml\n"
                "   Pero existe config.yaml.example (puedes renombrarlo)"
            )
        else:
            errores.append(
                "❌ No se encuentra ni config.yaml ni config.yaml.example\n"
                "   Asegúrate de tener los archivos de configuración"
            )

    # Verificar directorio de datos
    if not os.path.exists("data"):
        errores.append(
            "⚠️ No existe el directorio 'data'\n"
            "   Crea el directorio y coloca tu archivo correos.xlsx"
        )
    elif not os.path.exists("data/correos.xlsx"):
        errores.append(
            "⚠️ No se encuentra data/correos.xlsx\n"
            "   Crea el archivo con los correos a verificar"
        )

    return errores


def iniciar_con_checks():
    """Inicia la GUI después de verificar prerrequisitos."""
    print("🚀 Iniciando Interfaz Gráfica de Verificación de Correos")
    print("="*60)

    # Verificar prerrequisitos
    errores = verificar_prerrequisitos()

    if errores:
        print("\n⚠️ Se encontraron los siguientes problemas:")
        for error in errores:
            print(f"   {error}")

        print("\n¿Deseas continuar de todos modos? (s/N): ", end="")
        respuesta = input().strip().lower()

        if respuesta not in ['s', 'si', 'sí', 'y', 'yes']:
            print("\n❌ Inicio cancelado. Resuelve los problemas e intenta nuevamente.")
            return

    print("\n✅ Iniciando interfaz gráfica...")

    try:
        # Importar e iniciar GUI
        from gui import main as gui_main
        gui_main()

    except ImportError as e:
        print(f"\n❌ Error al importar módulos de la GUI: {e}")
        print("   Asegúrate de tener todos los archivos necesarios:")
        print("   - gui.py")
        print("   - gui_config_manager.py")
        print("   - gui_runner.py")
        sys.exit(1)

    except Exception as e:
        print(f"\n❌ Error al iniciar la GUI: {e}")
        sys.exit(1)


def mostrar_ayuda():
    """Muestra la ayuda del script."""
    print("""
📋 Interfaz Gráfica para Verificación de Correos OWA

Uso:
    python iniciar_gui.py          # Inicia GUI con verificación
    python iniciar_gui.py --help   # Muestra esta ayuda
    python gui.py                  # Inicia GUI directamente

Prerrequisitos:
    1. Archivo de sesión: state.json (ejecutar python copiar_sesion.py)
    2. Configuración: config.yaml (copiar desde config.yaml.example)
    3. Datos: data/correos.xlsx con correos a verificar

Estructura recomendada:
    verificacion-correo/
    ├── gui.py                    # Interfaz gráfica principal
    ├── iniciar_gui.py            # Script de lanzamiento
    ├── config.yaml               # Configuración
    ├── state.json                # Sesión guardada
    ├── data/
    │   └── correos.xlsx          # Correos a procesar
    └── ... (otros archivos del proyecto)

Para más información, consulta README_GUI.md
    """)


def main():
    """Función principal."""
    # Verificar argumentos de línea de comandos
    if len(sys.argv) > 1 and sys.argv[1] in ['-h', '--help', 'help']:
        mostrar_ayuda()
        return

    # Iniciar GUI con verificación
    iniciar_con_checks()


if __name__ == "__main__":
    main()