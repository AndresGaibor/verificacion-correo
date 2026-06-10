"""
Configuración centralizada para el proyecto de verificación de correos.
Lee la configuración desde config.yaml y crea el archivo si no existe.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from windows_compat import setup_console_encoding
setup_console_encoding()

import re
import shutil
import yaml


def _load_config():
    """
    Carga la configuración desde config.yaml.
    Si no existe, lo crea desde config.yaml.example
    """
    config_file = "config.yaml"
    example_file = "config.yaml.example"

    # Si no existe config.yaml, crearlo desde el example
    if not os.path.exists(config_file):
        if os.path.exists(example_file):
            print(f"[!] No se encontro {config_file}")
            print(f"[+] Creando {config_file} desde {example_file}")
            shutil.copy(example_file, config_file)
            print(f"[!] IMPORTANTE: Edita {config_file} con tus valores reales antes de continuar")
            print()
        else:
            raise FileNotFoundError(
                f"No se encontro ni {config_file} ni {example_file}. "
                "Asegurate de tener config.yaml.example en el directorio."
            )

    # Cargar el archivo YAML
    with open(config_file, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)

    return config


# Cargar configuración
_config = _load_config()

# ============================================================================
# URLS
# ============================================================================
PAGE_URL = _config['page_url']

# ============================================================================
# EMAILS POR DEFECTO (fallback si no se encuentra archivo Excel)
# ============================================================================
DEFAULT_EMAILS = _config['default_emails']

# ============================================================================
# PATRONES REGEX para extracción de datos
# ============================================================================
EMAIL_RE = re.compile(r'[\w.+-]+@[\w.-]+\.[a-z]{2,}', re.I)
PHONE_RE = re.compile(r'\b\d{6,}\b')  # 6+ dígitos
POSTAL_ADDR_RE = re.compile(r'\d{5}\s+[A-ZÁÉÍÓÚÑ\-\s]+', re.I)  # Ej: '12345 CIUDAD'
SIP_RE = re.compile(r'sip:[\w.+-]+@[\w.-]+', re.I)
NAME_RE = re.compile(r'([A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúñ\-\.\s]+,\s*[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúñ\-\s]+)')  # "APELLIDO, NOMBRE"

# ============================================================================
# SELECTORES CSS/XPATH para elementos de la interfaz OWA
# ============================================================================
SELECTORS = _config['selectors']

# ============================================================================
# TIEMPOS DE ESPERA (en milisegundos)
# ============================================================================
WAIT_TIMES = _config['wait_times']

# ============================================================================
# CONFIGURACIÓN DE NAVEGADOR
# ============================================================================
BROWSER_CONFIG = _config['browser']

# ============================================================================
# CONFIGURACIÓN DE EXCEL
# ============================================================================
EXCEL_CONFIG = _config['excel']

# ============================================================================
# CONFIGURACIÓN DE PROCESAMIENTO
# ============================================================================
PROCESSING_CONFIG = _config['processing']
