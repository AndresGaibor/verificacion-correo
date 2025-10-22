"""
Core functionality for email verification and contact extraction.

This module contains the main automation logic:
- config: Configuration management
- browser: Browser automation with Playwright
- extractor: Contact information extraction
- excel: Excel file reading/writing
- session: Browser session management
"""

from .config import Config
from .browser import BrowserAutomation
from .extractor import ContactExtractor
from .excel import ExcelReader, ExcelWriter
from .session import SessionManager

__all__ = [
    "Config",
    "BrowserAutomation",
    "ContactExtractor",
    "ExcelReader",
    "ExcelWriter",
    "SessionManager",
]