#!/usr/bin/env python3
"""
Gestor de sesiones para la interfaz gráfica.
Permite crear y gestionar sesiones de OWA sin usar la terminal.
"""

import threading
import time
import os
from tkinter import messagebox
from playwright.sync_api import sync_playwright
from config import PAGE_URL, BROWSER_CONFIG


class GUISessionManager:
    """Gestiona la creación y validación de sesiones de OWA desde la GUI."""

    def __init__(self, log_callback=None, status_callback=None):
        """
        Inicializa el gestor de sesiones.

        Args:
            log_callback: Función para enviar mensajes de log a la GUI
            status_callback: Función para actualizar el estado en la GUI
        """
        self.log_callback = log_callback
        self.status_callback = status_callback
        self.is_creating_session = False
        self.session_thread = None
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

    def create_session_interactive(self, session_file="state.json"):
        """
        Inicia la creación interactiva de una sesión.

        Args:
            session_file: Ruta donde guardar el archivo de sesión
        """
        if self.is_creating_session:
            self._safe_log("⚠️ Ya hay una sesión en creación")
            return

        self.is_creating_session = True
        self.session_thread = threading.Thread(
            target=self._create_session_thread,
            args=(session_file,),
            daemon=True
        )
        self.session_thread.start()

    def _create_session_thread(self, session_file):
        """
        Thread principal para crear la sesión.

        Args:
            session_file: Ruta donde guardar el archivo de sesión
        """
        try:
            self._safe_log("🚀 Iniciando creación de sesión...")
            self._safe_status("Abriendo navegador para autenticación...")

            # Iniciar Playwright
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(headless=False)
            self.context = self.browser.new_context()
            self.page = self.context.new_page()

            # Navegar a OWA
            self._safe_log(f"🌐 Navegando a {PAGE_URL}")
            self.page.goto(PAGE_URL)

            self._safe_log("✅ Navegador abierto. Por favor:")
            self._safe_log("   1. Inicia sesión en OWA manualmente")
            self._safe_log("   2. Espera a que la página principal cargue completamente")
            self._safe_log("   3. Cuando estés autenticado, haz clic en 'Guardar Sesión'")

            # Actualizar estado
            self._safe_status("Esperando autenticación manual...")

            # Esperar a que el usuario confirme
            self._wait_for_user_confirmation()

        except Exception as e:
            self._safe_log(f"❌ Error creando sesión: {str(e)}")
            self._safe_status("Error en creación de sesión")
            messagebox.showerror("Error", f"No se pudo crear la sesión: {str(e)}")
        finally:
            self._cleanup()

    def _wait_for_user_confirmation(self):
        """
        Espera la confirmación del usuario para guardar la sesión.
        Este método debe ser implementado en la GUI principal.
        """
        # La GUI principal implementará la lógica de confirmación
        # Este método es un placeholder
        pass

    def save_session_now(self, session_file="state.json"):
        """
        Guarda la sesión actual en el archivo especificado.

        Args:
            session_file: Ruta donde guardar el archivo de sesión

        Returns:
            bool: True si se guardó correctamente, False en caso contrario
        """
        try:
            if not self.context:
                self._safe_log("❌ No hay contexto de navegador activo")
                return False

            self._safe_log("💾 Guardando sesión...")
            self._safe_status("Guardando sesión...")

            # Guardar estado de la sesión
            self.context.storage_state(path=session_file)

            self._safe_log(f"✅ Sesión guardada en {session_file}")
            self._safe_status("Sesión guardada exitosamente")

            return True

        except Exception as e:
            self._safe_log(f"❌ Error guardando sesión: {str(e)}")
            self._safe_status("Error guardando sesión")
            return False

    def validate_session(self, session_file="state.json"):
        """
        Valida si una sesión existente es válida.

        Args:
            session_file: Ruta del archivo de sesión a validar

        Returns:
            tuple: (is_valid, error_message)
        """
        if not os.path.exists(session_file):
            return False, "El archivo de sesión no existe"

        try:
            self._safe_log("🔍 Validando sesión existente...")
            self._safe_status("Validando sesión...")

            # Intentar cargar la sesión
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(headless=True)
            self.context = self.browser.new_context(storage_state=session_file)
            self.page = self.context.new_page()

            # Navegar a OWA para verificar si la sesión es válida
            self.page.goto(PAGE_URL)
            self.page.wait_for_load_state("networkidle", timeout=10000)

            # Verificar si estamos en la página de login (sesión inválida)
            # O si estamos en el inbox (sesión válida)
            current_url = self.page.url
            page_title = self.page.title()

            # Si estamos en la página de login, la sesión no es válida
            if "login" in current_url.lower() or "signin" in current_url.lower():
                self._safe_log("❌ Sesión inválida: redirigido a página de login")
                return False, "La sesión ha expirado"

            # Si estamos en OWA principal, la sesión es válida
            if "owa" in current_url.lower() or "outlook" in current_url.lower():
                self._safe_log("✅ Sesión válida y activa")
                return True, ""

            # Verificación adicional por título de página
            if "outlook" in page_title.lower() or "correo" in page_title.lower():
                self._safe_log("✅ Sesión válida y activa")
                return True, ""

            self._safe_log("⚠️ No se pudo determinar el estado de la sesión")
            return False, "Estado de sesión desconocido"

        except Exception as e:
            self._safe_log(f"❌ Error validando sesión: {str(e)}")
            return False, f"Error validando sesión: {str(e)}"
        finally:
            self._cleanup()

    def cancel_session_creation(self):
        """Cancela el proceso de creación de sesión."""
        if self.is_creating_session:
            self._safe_log("🛑 Cancelando creación de sesión...")
            self.is_creating_session = False
            self._cleanup()

    def _cleanup(self):
        """Limpia los recursos del navegador."""
        try:
            if self.page:
                self.page.close()
            if self.context:
                self.context.close()
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()
        except:
            pass  # Ignorar errores en cleanup
        finally:
            self.page = None
            self.context = None
            self.browser = None
            self.playwright = None
            self.is_creating_session = False

    def _safe_log(self, message):
        """Envía mensaje de log de forma segura."""
        if self.log_callback:
            self.log_callback(message)

    def _safe_status(self, status):
        """Actualiza el estado de forma segura."""
        if self.status_callback:
            self.status_callback(status)

    def get_session_status(self, session_file="state.json"):
        """
        Obtiene el estado actual de la sesión.

        Args:
            session_file: Ruta del archivo de sesión

        Returns:
            dict: Información sobre el estado de la sesión
        """
        if not os.path.exists(session_file):
            return {
                "exists": False,
                "valid": False,
                "status": "No existe",
                "message": "El archivo de sesión no existe"
            }

        # Verificar fecha de modificación
        try:
            file_mtime = os.path.getmtime(session_file)
            import datetime
            file_age = datetime.datetime.now() - datetime.datetime.fromtimestamp(file_mtime)
            age_hours = file_age.total_seconds() / 3600

            if age_hours > 24:
                age_str = f"Hace {int(age_hours)} horas"
                status_color = "orange"
            else:
                age_str = f"Hace {int(age_hours)}h"
                status_color = "green"

            return {
                "exists": True,
                "valid": None,  # Se necesita validación
                "age_hours": age_hours,
                "age_str": age_str,
                "status": "Requiere validación",
                "status_color": status_color,
                "message": f"Sesión guardada {age_str}"
            }
        except Exception as e:
            return {
                "exists": True,
                "valid": False,
                "status": "Error",
                "message": f"Error verificando sesión: {str(e)}"
            }