"""
Gestor de configuración para la interfaz gráfica.
Permite cargar, modificar y guardar la configuración desde la GUI.
"""

import os
import yaml
from tkinter import messagebox


class GUIConfigManager:
    """Gestiona la configuración de la aplicación desde la interfaz gráfica."""

    def __init__(self):
        self.config_file = "config.yaml"
        self.config = self._load_config()

    def _load_config(self):
        """Carga la configuración desde config.yaml."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return yaml.safe_load(f)
            else:
                # Si no existe, crear configuración por defecto
                return self._get_default_config()
        except Exception as e:
            messagebox.showerror("Error de Configuración",
                                f"No se pudo cargar la configuración:\n{str(e)}")
            return self._get_default_config()

    def _get_default_config(self):
        """Retorna configuración por defecto."""
        return {
            'page_url': 'https://correoweb.madrid.org/owa/#path=/mail',
            'default_emails': ['user1@madrid.org', 'user2@madrid.org'],
            'browser': {
                'headless': False,
                'session_file': 'state.json'
            },
            'excel': {
                'default_file': 'data/correos.xlsx',
                'start_row': 2,
                'email_column': 1
            },
            'selectors': {
                'new_message_btn': 'button[title="Escribir un mensaje nuevo (N)"]',
                'to_field_role': 'textbox',
                'to_field_name': 'Para',
                'popup': 'div._pe_Y[ispopup="1"]',
                'discard_btn': 'button[aria-label="Descartar"]'
            },
            'wait_times': {
                'after_new_message': 1000,
                'after_fill_to': 3000,
                'after_blur': 500,
                'popup_visible': 5000,
                'after_click_token': 2000,
                'popup_load_data': 5000,
                'after_close_popup': 1000,
                'before_discard': 2000
            },
            'processing': {
                'batch_size': 10
            }
        }

    def get_config_value(self, key_path, default=None):
        """
        Obtiene un valor de configuración usando una ruta de clave.

        Args:
            key_path: Ruta de la clave (ej: 'page_url' o 'processing.batch_size')
            default: Valor por defecto si no se encuentra

        Returns:
            Valor de configuración o default
        """
        keys = key_path.split('.')
        value = self.config

        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            return default

    def set_config_value(self, key_path, value):
        """
        Establece un valor de configuración usando una ruta de clave.

        Args:
            key_path: Ruta de la clave (ej: 'page_url' o 'processing.batch_size')
            value: Nuevo valor
        """
        keys = key_path.split('.')
        config_section = self.config

        # Navegar hasta la sección padre
        for key in keys[:-1]:
            if key not in config_section:
                config_section[key] = {}
            config_section = config_section[key]

        # Establecer el valor
        config_section[keys[-1]] = value

    def save_config(self):
        """Guarda la configuración actual en config.yaml."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                yaml.dump(self.config, f, default_flow_style=False,
                         allow_unicode=True, sort_keys=False)
            return True
        except Exception as e:
            messagebox.showerror("Error al Guardar",
                               f"No se pudo guardar la configuración:\n{str(e)}")
            return False

    def reload_config(self):
        """Recarga la configuración desde el archivo."""
        self.config = self._load_config()

    def validate_url(self, url):
        """Valida que la URL tenga un formato correcto."""
        if not url or not url.strip():
            return False, "La URL no puede estar vacía"

        url = url.strip()
        if not (url.startswith('http://') or url.startswith('https://')):
            return False, "La URL debe comenzar con http:// o https://"

        return True, ""

    def validate_batch_size(self, batch_size):
        """Valida que el tamaño de lote sea un número entero positivo."""
        try:
            size = int(batch_size)
            if size <= 0:
                return False, "El tamaño de lote debe ser mayor que 0"
            if size > 50:
                return False, "El tamaño de lote no debe exceder 50"
            return True, ""
        except ValueError:
            return False, "El tamaño de lote debe ser un número entero"

    def get_current_settings(self):
        """Retorna la configuración actual como diccionario plano."""
        return {
            'page_url': self.get_config_value('page_url'),
            'batch_size': self.get_config_value('processing.batch_size', 10),
            'headless': self.get_config_value('browser.headless', False),
            'excel_file': self.get_config_value('excel.default_file')
        }

    def update_settings(self, settings):
        """
        Actualiza múltiples valores de configuración.

        Args:
            settings: Diccionario con los valores a actualizar

        Returns:
            tuple: (success, error_message)
        """
        try:
            # Validar URL
            if 'page_url' in settings:
                valid, msg = self.validate_url(settings['page_url'])
                if not valid:
                    return False, msg
                self.set_config_value('page_url', settings['page_url'].strip())

            # Validar batch_size
            if 'batch_size' in settings:
                valid, msg = self.validate_batch_size(settings['batch_size'])
                if not valid:
                    return False, msg
                self.set_config_value('processing.batch_size', int(settings['batch_size']))

            # Actualizar otros valores
            if 'headless' in settings:
                self.set_config_value('browser.headless', bool(settings['headless']))

            if 'excel_file' in settings:
                self.set_config_value('excel.default_file', settings['excel_file'].strip())

            return True, ""

        except Exception as e:
            return False, f"Error actualizando configuración: {str(e)}"