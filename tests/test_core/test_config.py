"""Tests for core.config module."""

import os
import json
import tempfile
from pathlib import Path
from unittest.mock import mock_open, patch, MagicMock, PropertyMock

import pytest
import yaml

from verificacion_correo.core.config import (
    Config,
    BrowserConfig,
    ExcelConfig,
    ProcessingConfig,
    Selectors,
    WaitTimes,
    RegexPatterns,
    get_config,
    reload_config,
)


SAMPLE_YAML = """
page_url: https://correoweb.madrid.org/owa/#path=/mail
default_emails:
  - ASP164@MADRID.ORG
  - AGM564@MADRID.ORG
browser:
  headless: false
  session_file: state.json
excel:
  default_file: data/correos.xlsx
  start_row: 2
  email_column: 1
processing:
  batch_size: 10
  discard_draft: false
"""


class TestConfig:
    def test_init_with_valid_yaml(self):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(SAMPLE_YAML)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            assert config.page_url == "https://correoweb.madrid.org/owa/#path=/mail"
            assert config.default_emails == ["ASP164@MADRID.ORG", "AGM564@MADRID.ORG"]
            assert config.processing.batch_size == 10
        finally:
            os.unlink(tmp_path)

    def test_init_with_missing_file_raises(self):
        with pytest.raises(FileNotFoundError):
            Config(config_path="/nonexistent/config.yaml")

    def test_init_with_invalid_yaml_raises(self):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write("{invalid: yaml: [}")
            tmp_path = f.name

        try:
            with pytest.raises(ValueError, match="Error en el formato YAML"):
                Config(config_path=tmp_path)
        finally:
            os.unlink(tmp_path)

    def test_init_with_empty_yaml_uses_defaults(self):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write("")
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            assert config.page_url == "https://correoweb.madrid.org/owa/#path=/mail"
            assert config.default_emails == []
            assert config.processing.batch_size == 10
            assert config.browser.headless is False
        finally:
            os.unlink(tmp_path)

    def test_init_with_partial_config(self):
        partial = """
page_url: https://custom.url/owa
default_emails:
  - test@example.com
"""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(partial)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            assert config.page_url == "https://custom.url/owa"
            assert config.default_emails == ["test@example.com"]
            # Defaults for unspecified fields
            assert config.processing.batch_size == 10
            assert config.browser.headless is False
        finally:
            os.unlink(tmp_path)

    def test_validate_clean_config(self):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(SAMPLE_YAML)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            issues = config.validate()
            assert issues == []
        finally:
            os.unlink(tmp_path)

    def test_validate_empty_page_url(self):
        yaml_content = "page_url: ''\ndefault_emails:\n  - test@example.com"
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(yaml_content)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            issues = config.validate()
            assert "page_url no puede estar vacío" in issues
        finally:
            os.unlink(tmp_path)

    def test_validate_empty_default_emails(self):
        yaml_content = "page_url: https://example.com\ndefault_emails: []"
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(yaml_content)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            issues = config.validate()
            assert "default_emails no puede estar vacío" in issues
        finally:
            os.unlink(tmp_path)

    def test_validate_zero_batch_size(self):
        yaml_content = (
            "page_url: https://example.com\n"
            "default_emails:\n  - test@example.com\n"
            "processing:\n  batch_size: 0\n"
        )
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(yaml_content)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            issues = config.validate()
            assert "batch_size debe ser mayor que 0" in issues
        finally:
            os.unlink(tmp_path)

    def test_validate_negative_batch_size(self):
        yaml_content = (
            "page_url: https://example.com\n"
            "default_emails:\n  - test@example.com\n"
            "processing:\n  batch_size: -1\n"
        )
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(yaml_content)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            issues = config.validate()
            assert "batch_size debe ser mayor que 0" in issues
        finally:
            os.unlink(tmp_path)

    def test_validate_zero_start_row(self):
        yaml_content = (
            "page_url: https://example.com\n"
            "default_emails:\n  - test@example.com\n"
            "excel:\n  start_row: 0\n"
        )
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(yaml_content)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            issues = config.validate()
            assert "start_row debe ser mayor o igual a 1" in issues
        finally:
            os.unlink(tmp_path)

    def test_validate_zero_email_column(self):
        yaml_content = (
            "page_url: https://example.com\n"
            "default_emails:\n  - test@example.com\n"
            "excel:\n  email_column: 0\n"
        )
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(yaml_content)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            issues = config.validate()
            assert "email_column debe ser mayor o igual a 1" in issues
        finally:
            os.unlink(tmp_path)

    def test_validate_multiple_issues(self):
        yaml_content = "page_url: ''\ndefault_emails: []"
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(yaml_content)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            issues = config.validate()
            assert len(issues) >= 2
        finally:
            os.unlink(tmp_path)

    def test_to_dict(self):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(SAMPLE_YAML)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            result = config.to_dict()

            assert result["page_url"] == "https://correoweb.madrid.org/owa/#path=/mail"
            assert result["default_emails"] == ["ASP164@MADRID.ORG", "AGM564@MADRID.ORG"]
            assert result["browser"]["headless"] is False
            assert result["browser"]["session_file"].endswith("state.json")
            assert result["excel"]["start_row"] == 2
            assert result["excel"]["email_column"] == 1
            assert result["processing"]["batch_size"] == 10
            assert result["selectors"]["new_message_btn"] == 'button[title="Escribir un mensaje nuevo (N)"]'
            assert result["wait_times"]["after_new_message"] == 1000
        finally:
            os.unlink(tmp_path)

    def test_save(self):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(SAMPLE_YAML)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            config.page_url = "https://modified.url/owa"
            config.save()

            with open(tmp_path) as f:
                saved = yaml.safe_load(f)
            assert saved["page_url"] == "https://modified.url/owa"
        finally:
            os.unlink(tmp_path)

    def test_get_session_file_path(self):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(SAMPLE_YAML)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            path = config.get_session_file_path()
            assert "state.json" in path
            assert os.path.isabs(path)
        finally:
            os.unlink(tmp_path)

    def test_get_excel_file_path(self):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(SAMPLE_YAML)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            path = config.get_excel_file_path()
            assert "correos.xlsx" in path
            assert os.path.isabs(path)
        finally:
            os.unlink(tmp_path)

    def test_ensure_data_directory(self, tmp_path):
        yaml_content = f"excel:\n  default_file: {tmp_path / 'sub' / 'data.xlsx'}"
        cfg_path = tmp_path / "cfg.yaml"
        cfg_path.write_text(yaml_content)

        config = Config(config_path=str(cfg_path))
        config.ensure_data_directory()
        assert (tmp_path / "sub").exists()

    def test_repr(self):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False) as f:
            f.write(SAMPLE_YAML)
            tmp_path = f.name

        try:
            config = Config(config_path=tmp_path)
            rep = repr(config)
            assert "Config(" in rep
            assert "page_url=" in rep
            assert "batch_size=" in rep
        finally:
            os.unlink(tmp_path)

    def test_absolute_paths_in_post_init(self):
        browser_cfg = BrowserConfig(headless=True, session_file="custom_state.json")
        assert os.path.isabs(browser_cfg.session_file)
        assert browser_cfg.headless is True

        excel_cfg = ExcelConfig(default_file="custom/data.xlsx")
        assert os.path.isabs(excel_cfg.default_file)

    def test_processing_config_discard_draft_default(self):
        cfg = ProcessingConfig()
        assert cfg.discard_draft is False

    def test_selectors_defaults(self):
        sel = Selectors()
        assert "Escribir un mensaje nuevo" in sel.new_message_btn
        assert sel.to_field_role == "textbox"

    def test_wait_times_defaults(self):
        wt = WaitTimes()
        assert wt.after_new_message == 1000
        assert wt.popup_visible == 5000
        assert wt.before_discard == 2000

    def test_regex_patterns_are_compiled(self):
        rp = RegexPatterns()
        assert rp.EMAIL.search("user@example.com") is not None
        assert rp.PHONE.search("123456789") is not None
        assert rp.POSTAL_ADDR.search("28001 Madrid") is not None
        assert rp.SIP.search("sip:user@domain.com") is not None
        assert rp.NAME.search("GARCIA, JUAN") is not None


class TestGetConfig:
    def setup_method(self):
        reload_config()

    def test_get_config_returns_instance(self):
        with patch(
            "verificacion_correo.core.config.Config._get_default_config_path",
            return_value="/tmp/test_config.yaml",
        ), patch(
            "verificacion_correo.core.config.Config._load_config",
            return_value={
                "page_url": "https://example.com",
                "default_emails": ["test@example.com"],
            },
        ):
            config = get_config()
            assert isinstance(config, Config)
            assert config.page_url == "https://example.com"

    def test_get_config_is_singleton(self):
        reload_config()
        with patch(
            "verificacion_correo.core.config.Config._get_default_config_path",
            return_value="/tmp/test_config.yaml",
        ), patch(
            "verificacion_correo.core.config.Config._load_config",
            return_value={
                "page_url": "https://example.com",
                "default_emails": ["test@example.com"],
            },
        ):
            config1 = get_config()
            config2 = get_config()
            assert config1 is config2

    def test_reload_config_resets_singleton(self):
        reload_config()
        with patch(
            "verificacion_correo.core.config.Config._get_default_config_path",
            return_value="/tmp/test_config.yaml",
        ), patch(
            "verificacion_correo.core.config.Config._load_config",
            return_value={
                "page_url": "https://example.com",
                "default_emails": ["test@example.com"],
            },
        ):
            config1 = get_config()
            reload_config()
            config2 = get_config()
            assert config1 is not config2

    def test_get_config_shows_warnings(self):
        reload_config()
        with patch(
            "verificacion_correo.core.config.Config._get_default_config_path",
            return_value="/tmp/test_config.yaml",
        ), patch(
            "verificacion_correo.core.config.Config._load_config",
            return_value={"page_url": "", "default_emails": []},
        ), patch("builtins.print") as mock_print:
            get_config()
            mock_print.assert_any_call("⚠️ Advertencias de configuración:")


class TestDefaultConfigCreation:
    def test_create_default_config_at_path(self, tmp_path):
        cfg_path = tmp_path / "generated.yaml"
        config = Config.__new__(Config)
        result_path = config._create_default_config_at_path(cfg_path)

        assert result_path == str(cfg_path)
        assert cfg_path.exists()

        with open(cfg_path) as f:
            data = yaml.safe_load(f)

        assert data["page_url"] == "https://correoweb.madrid.org/owa/#path=/mail"
        assert data["default_emails"] == ["ASP164@MADRID.ORG", "AGM564@MADRID.ORG"]

    def test_create_default_config(self, tmp_path):
        old_cwd = Path.cwd()
        os.chdir(tmp_path)
        try:
            config = Config.__new__(Config)
            result_path = config._create_default_config()
            assert Path("config.yaml").exists()
            assert result_path == str(Path("config.yaml"))
        finally:
            os.chdir(old_cwd)

    def test_create_default_config_at_path_creates_parent(self, tmp_path):
        cfg_path = tmp_path / "nested" / "cfg.yaml"
        config = Config.__new__(Config)
        result_path = config._create_default_config_at_path(cfg_path)
        assert cfg_path.exists()
        assert result_path == str(cfg_path)

    def test_get_default_config_path_uses_new_config_first(self, tmp_path):
        old_cwd = Path.cwd()
        os.chdir(tmp_path)
        try:
            new_cfg = tmp_path / "config" / "default.yaml"
            new_cfg.parent.mkdir(exist_ok=True)
            new_cfg.write_text("page_url: https://new.example.com\n")

            config = Config.__new__(Config)
            path = config._get_default_config_path()
            assert "config/default.yaml" in path
        finally:
            os.chdir(old_cwd)

    def test_get_default_config_path_falls_back_to_legacy(self, tmp_path):
        old_cwd = Path.cwd()
        os.chdir(tmp_path)
        try:
            legacy = tmp_path / "config.yaml"
            legacy.write_text("page_url: https://legacy.example.com\n")

            config = Config.__new__(Config)
            path = config._get_default_config_path()
            assert "config.yaml" in path
        finally:
            os.chdir(old_cwd)

    def test_get_default_config_path_uses_example(self, tmp_path):
        old_cwd = Path.cwd()
        os.chdir(tmp_path)
        try:
            example = tmp_path / "config.yaml.example"
            example.write_text("page_url: https://example.com\n")

            config = Config.__new__(Config)
            path = config._get_default_config_path()
            assert "config.yaml" in path
            assert Path("config.yaml").exists()
        finally:
            os.chdir(old_cwd)

    def test_get_default_config_path_frozen_with_meipass(self):
        with patch("sys.frozen", True, create=True), patch(
            "sys._MEIPASS", "/bundle", create=True
        ), patch("sys.executable", "/usr/local/bin/app"), patch(
            "os.path.dirname", return_value="/usr/local/bin"
        ):
            with tempfile.NamedTemporaryFile(suffix=".yaml", delete=False) as f:
                f.write(b"page_url: https://bundle.example.com\n")
                bundle_cfg = Path("/usr/local/bin/config.yaml")
                # We mock the exists check
                with patch.object(Path, "exists") as mock_exists:
                    mock_exists.side_effect = lambda: True
                    config = Config.__new__(Config)
                    with patch.object(config, "_setup_config_for_executable") as mock_setup:
                        mock_setup.return_value = str(bundle_cfg)
                        path = config._get_default_config_path()
                        assert path == str(bundle_cfg)
