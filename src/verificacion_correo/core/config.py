"""
Configuration management for verificacion-correo.

This module handles loading and managing configuration from YAML files,
with proper fallbacks and validation.
"""

import os
import re
import shutil
import sys
from pathlib import Path
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
import yaml


@dataclass
class BrowserConfig:
    """Browser configuration settings."""
    headless: bool = False
    session_file: str = "state.json"

    def __post_init__(self):
        # Ensure session file path is absolute
        if not os.path.isabs(self.session_file):
            self.session_file = os.path.abspath(self.session_file)


@dataclass
class ExcelConfig:
    """Excel file configuration."""
    default_file: str = "data/correos.xlsx"
    start_row: int = 2
    email_column: int = 1  # Column A

    def __post_init__(self):
        # Ensure file path is absolute
        if not os.path.isabs(self.default_file):
            self.default_file = os.path.abspath(self.default_file)


@dataclass
class ProcessingConfig:
    """Processing configuration."""
    batch_size: int = 10
    discard_draft: bool = True  # Discard draft after processing by default


@dataclass
class Selectors:
    """CSS selectors for OWA interface elements."""
    new_message_btn: str = 'button[title="Escribir un mensaje nuevo (N)"]'
    to_field_role: str = "textbox"
    to_field_name: str = "Para"
    popup: str = "div._pe_Y[ispopup='1']"
    discard_btn: str = 'button[aria-label="Descartar"]'


@dataclass
class WaitTimes:
    """Wait time configurations (in milliseconds)."""
    after_new_message: int = 1000
    after_fill_to: int = 3000
    after_blur: int = 500
    popup_visible: int = 5000
    after_click_token: int = 2000
    popup_load_data: int = 5000
    after_close_popup: int = 1000
    before_discard: int = 2000


@dataclass
class RegexPatterns:
    """Regular expression patterns for data extraction."""
    EMAIL: re.Pattern = re.compile(r'[\w.+-]+@[\w.-]+\.[a-z]{2,}', re.I)
    PHONE: re.Pattern = re.compile(r'\b\d{6,}\b')  # 6+ digits
    POSTAL_ADDR: re.Pattern = re.compile(r'\d{5}\s+[A-ZÁÉÍÓÚÑ\-\s]+', re.I)
    SIP: re.Pattern = re.compile(r'sip:[\w.+-]+@[\w.-]+', re.I)
    NAME: re.Pattern = re.compile(
        r'([A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúñ\-\.\s]+,\s*[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúñ\-\s]+)'
    )  # "APELLIDO, NOMBRE"


class Config:
    """
    Main configuration class that loads and manages all settings.

    Attributes:
        page_url: OWA page URL
        default_emails: Fallback email list
        browser: Browser configuration
        excel: Excel configuration
        processing: Processing configuration
        selectors: CSS selectors
        wait_times: Wait times
        patterns: Regex patterns
    """

    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize configuration.

        Args:
            config_path: Path to config file. If None, uses default location.
        """
        self._config_path = config_path or self._get_default_config_path()
        self._config_data = self._load_config()
        self._initialize_components()

    def _get_default_config_path(self) -> str:
        """Get default configuration file path."""
        # First, try to create from bundled resources if we're in a PyInstaller executable
        if getattr(sys, 'frozen', False):
            return self._setup_config_for_executable()

        # Try config/default.yaml first (new structure)
        new_config = Path("config/default.yaml")
        if new_config.exists():
            return str(new_config)

        # Fall back to root config.yaml (legacy)
        legacy_config = Path("config.yaml")
        if legacy_config.exists():
            return str(legacy_config)

        # Create from example if available
        example_config = Path("config.yaml.example")
        if example_config.exists():
            print(f"⚠ No se encontró archivo de configuración")
            print(f"✓ Creando config.yaml desde config.yaml.example")
            shutil.copy(example_config, "config.yaml")
            print(f"⚠ IMPORTANTE: Edita config.yaml con tus valores reales")
            return "config.yaml"

        # Create default config if nothing exists
        return self._create_default_config()

    def _setup_config_for_executable(self) -> str:
        """Setup configuration for PyInstaller executable."""
        # Get the directory where the executable is located
        if hasattr(sys, '_MEIPASS'):
            # We're running in a PyInstaller bundle
            bundle_dir = Path(sys._MEIPASS)
            exec_dir = Path(os.path.dirname(sys.executable))
        else:
            # We're running in development
            exec_dir = Path.cwd()

        config_path = exec_dir / "config.yaml"

        # If config doesn't exist, try to copy from bundle
        if not config_path.exists():
            example_in_bundle = None
            if hasattr(sys, '_MEIPASS'):
                example_in_bundle = bundle_dir / "config.yaml.example"

            if example_in_bundle and example_in_bundle.exists():
                print("✓ Creando config.yaml desde recursos del ejecutable")
                shutil.copy(example_in_bundle, config_path)
            else:
                # Create minimal default config
                self._create_default_config_at_path(config_path)

        return str(config_path)

    def _create_default_config(self) -> str:
        """Create default configuration file."""
        config_path = Path("config.yaml")
        return self._create_default_config_at_path(config_path)

    def _create_default_config_at_path(self, config_path: Path) -> str:
        """Create default configuration file at specific path."""
        default_config = {
            'page_url': 'https://correoweb.madrid.org/owa/#path=/mail',
            'default_emails': [
                'ASP164@MADRID.ORG',
                'AGM564@MADRID.ORG'
            ],
            'browser': {
                'headless': False,
                'session_file': 'state.json'
            },
            'excel': {
                'default_file': 'data/correos.xlsx',
                'start_row': 2,
                'email_column': 1
            },
            'processing': {
                'batch_size': 10,
                'discard_draft': True
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
            }
        }

        # Ensure parent directory exists
        config_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.dump(default_config, f, default_flow_style=False, allow_unicode=True)

            print(f"✓ Configuración por defecto creada en: {config_path}")
            print("⚠ IMPORTANTE: Revisa y ajusta la configuración según necesites")
            return str(config_path)

        except Exception as e:
            raise FileNotFoundError(f"No se pudo crear configuración por defecto: {e}")

    def _load_config(self) -> Dict[str, Any]:
        """Load configuration from YAML file."""
        try:
            with open(self._config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}
        except FileNotFoundError:
            raise FileNotFoundError(
                f"No se encuentra el archivo de configuración: {self._config_path}"
            )
        except yaml.YAMLError as e:
            raise ValueError(f"Error en el formato YAML de {self._config_path}: {e}")

    def _initialize_components(self):
        """Initialize configuration components from loaded data."""
        self.page_url: str = self._config_data.get('page_url', 'https://correoweb.madrid.org/owa/#path=/mail')
        self.default_emails: List[str] = self._config_data.get('default_emails', [])

        # Initialize structured configs with fallbacks
        browser_data = self._config_data.get('browser', {})
        self.browser = BrowserConfig(**browser_data)

        excel_data = self._config_data.get('excel', {})
        self.excel = ExcelConfig(**excel_data)

        processing_data = self._config_data.get('processing', {})
        self.processing = ProcessingConfig(**processing_data)

        selectors_data = self._config_data.get('selectors', {})
        self.selectors = Selectors(**selectors_data)

        wait_times_data = self._config_data.get('wait_times', {})
        self.wait_times = WaitTimes(**wait_times_data)

        # Regex patterns are fixed
        self.patterns = RegexPatterns()

    def ensure_data_directory(self):
        """Ensure data directory exists."""
        data_dir = Path(self.excel.default_file).parent
        data_dir.mkdir(parents=True, exist_ok=True)

    def get_session_file_path(self) -> str:
        """Get absolute path to session file."""
        return self.browser.session_file

    def get_excel_file_path(self) -> str:
        """Get absolute path to Excel file."""
        return self.excel.default_file

    def validate(self) -> List[str]:
        """
        Validate configuration and return list of issues.

        Returns:
            List of validation error messages. Empty if valid.
        """
        issues = []

        if not self.page_url:
            issues.append("page_url no puede estar vacío")

        if not self.default_emails:
            issues.append("default_emails no puede estar vacío")

        if self.processing.batch_size <= 0:
            issues.append("batch_size debe ser mayor que 0")

        if self.excel.start_row < 1:
            issues.append("start_row debe ser mayor o igual a 1")

        if self.excel.email_column < 1:
            issues.append("email_column debe ser mayor o igual a 1")

        return issues

    def __repr__(self) -> str:
        """String representation of config."""
        return f"Config(page_url={self.page_url!r}, batch_size={self.processing.batch_size})"


# Global configuration instance
_config: Optional[Config] = None


def get_config() -> Config:
    """
    Get global configuration instance.

    Returns:
        Global Config instance (lazy-loaded).
    """
    global _config
    if _config is None:
        _config = Config()
        # Validate and report issues
        issues = _config.validate()
        if issues:
            print("⚠️ Advertencias de configuración:")
            for issue in issues:
                print(f"  - {issue}")
            print()

    return _config


def reload_config():
    """Reload global configuration instance."""
    global _config
    _config = None