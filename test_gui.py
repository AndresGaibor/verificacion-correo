#!/usr/bin/env python3
"""
Script de prueba para la interfaz gráfica.
Inicia la GUI sin ejecutar procesamiento real.
"""

import tkinter as tk
from tkinter import messagebox
import os
import sys

def test_gui_imports():
    """Prueba las importaciones necesarias para la GUI."""
    try:
        from gui_config_manager import GUIConfigManager
        from gui_runner import GUIRunner
        from gui import VerificacionCorreosGUI
        return True, "Todas las importaciones exitosas"
    except ImportError as e:
        return False, f"Error de importación: {e}"
    except Exception as e:
        return False, f"Error inesperado: {e}"

def test_gui_components():
    """Prueba los componentes individuales de la GUI."""
    results = []

    # Probar GUIConfigManager
    try:
        from gui_config_manager import GUIConfigManager
        config_manager = GUIConfigManager()
        settings = config_manager.get_current_settings()
        results.append("✅ GUIConfigManager: Funcionando")
    except Exception as e:
        results.append(f"❌ GUIConfigManager: {e}")

    # Probar GUIRunner
    try:
        from gui_runner import GUIRunner
        runner = GUIRunner()
        stats = runner.get_processing_stats()
        results.append("✅ GUIRunner: Funcionando")
    except Exception as e:
        results.append(f"❌ GUIRunner: {e}")

    return results

def main():
    """Función principal de prueba."""
    print("🧪 PRUEBA DE INTERFAZ GRÁFICA")
    print("="*40)

    # 1. Probar importaciones
    print("\n1. Probando importaciones...")
    success, message = test_gui_imports()
    if success:
        print(f"✅ {message}")
    else:
        print(f"❌ {message}")
        return

    # 2. Probar componentes
    print("\n2. Probando componentes...")
    results = test_gui_components()
    for result in results:
        print(f"   {result}")

    # 3. Iniciar GUI (si todo está bien)
    print("\n3. Iniciando interfaz gráfica de prueba...")
    print("   (La ventana se cerrará automáticamente después de 5 segundos)")

    try:
        root = tk.Tk()
        from gui import VerificacionCorreosGUI
        app = VerificacionCorreosGUI(root)

        # Configurar para cerrar automáticamente después de 5 segundos
        def auto_close():
            root.destroy()
            print("\n✅ Prueba completada exitosamente")
            print("   La interfaz gráfica funciona correctamente")

        root.after(5000, auto_close)
        root.mainloop()

    except Exception as e:
        print(f"\n❌ Error al iniciar GUI: {e}")

if __name__ == "__main__":
    main()