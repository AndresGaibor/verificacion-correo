"""
Core functionality for email verification and contact extraction.

This module contains the main automation logic:
- config: Configuration management
- browser: Browser automation with Playwright
- extractor: Contact information extraction
- excel: Excel file reading/writing
- session: Browser session management
"""

from verificacion_correo.core.config import Config
from verificacion_correo.core.browser import BrowserAutomation
from verificacion_correo.core.extractor import ContactExtractor
from verificacion_correo.core.excel import ExcelReader, ExcelWriter
from verificacion_correo.core.session import SessionManager
from verificacion_correo.core.gal_scraper import scrape_gal
from verificacion_correo.core.platform import open_file, open_folder, open_in_explorer

__all__ = [
    "Config",
    "BrowserAutomation",
    "ContactExtractor",
    "ExcelReader",
    "ExcelWriter",
    "SessionManager",
    "scrape_gal",
    "open_file",
    "open_folder",
    "open_in_explorer",
]