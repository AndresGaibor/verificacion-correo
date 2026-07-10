"""Tests for core.__init__ module."""

from verificacion_correo import __version__, __author__, __email__


class TestPackageVersion:
    def test_version_is_string(self):
        assert isinstance(__version__, str)

    def test_version_is_correct(self):
        assert __version__ == "2.0.0"

    def test_version_format(self):
        parts = __version__.split(".")
        assert len(parts) == 3
        major, minor, patch = parts
        assert major.isdigit()
        assert minor.isdigit()
        assert patch.isdigit()


class TestPackageAuthor:
    def test_author_is_string(self):
        assert isinstance(__author__, str)

    def test_author_is_set(self):
        assert len(__author__) > 0


class TestPackageEmail:
    def test_email_is_string(self):
        assert isinstance(__email__, str)

    def test_email_is_set(self):
        assert len(__email__) > 0


class TestCoreImports:
    def test_config_is_importable(self):
        from verificacion_correo.core import Config
        assert Config is not None

    def test_browser_automation_is_importable(self):
        from verificacion_correo.core import BrowserAutomation
        assert BrowserAutomation is not None

    def test_contact_extractor_is_importable(self):
        from verificacion_correo.core import ContactExtractor
        assert ContactExtractor is not None

    def test_excel_reader_is_importable(self):
        from verificacion_correo.core import ExcelReader
        assert ExcelReader is not None

    def test_excel_writer_is_importable(self):
        from verificacion_correo.core import ExcelWriter
        assert ExcelWriter is not None

    def test_session_manager_is_importable(self):
        from verificacion_correo.core import SessionManager
        assert SessionManager is not None

    def test_scrape_gal_is_importable(self):
        from verificacion_correo.core import scrape_gal
        assert scrape_gal is not None

    def test_all_exports(self):
        from verificacion_correo.core import __all__
        expected = [
            "Config",
            "BrowserAutomation",
            "ContactExtractor",
            "ExcelReader",
            "ExcelWriter",
            "SessionManager",
            "scrape_gal",
        ]
        for item in expected:
            assert item in __all__


class TestUtilsImports:
    def test_setup_logging_is_importable(self):
        from verificacion_correo.utils import setup_logging
        assert setup_logging is not None

    def test_get_logger_is_importable(self):
        from verificacion_correo.utils import get_logger
        assert get_logger is not None


class TestPackageAll:
    def test_package_all_contains_expected(self):
        from verificacion_correo import __all__
        expected = [
            "config",
            "browser",
            "extractor",
            "excel",
            "session",
            "cli_main",
            "gui_main",
        ]
        for item in expected:
            assert item in __all__
