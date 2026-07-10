"""Tests for utils.logging module."""

import logging
import os
import tempfile
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

from verificacion_correo.utils.logging import (
    setup_logging,
    get_logger,
    create_default_logging_config,
)


class TestGetLogger:
    def test_returns_logger_instance(self):
        logger = get_logger("test_logger")
        assert isinstance(logger, logging.Logger)

    def test_returns_same_logger_for_same_name(self):
        logger1 = get_logger("same_name")
        logger2 = get_logger("same_name")
        assert logger1 is logger2

    def test_different_names_different_loggers(self):
        logger1 = get_logger("unique_name_1")
        logger2 = get_logger("unique_name_2")
        assert logger1 is not logger2

    def test_logger_name_is_correct(self):
        logger = get_logger("my_custom_name")
        assert logger.name == "my_custom_name"

    def test_can_log_messages(self, caplog):
        logger = get_logger("test_logger_can_log")
        with caplog.at_level(logging.DEBUG):
            logger.info("Test info message")
            logger.debug("Test debug message")
            logger.warning("Test warning message")
        assert len(caplog.records) >= 1


class TestSetupLogging:
    def test_basic_setup(self):
        setup_logging(level="DEBUG")
        root_logger = logging.getLogger()
        assert root_logger.level == logging.DEBUG

    def test_setup_with_log_file(self, tmp_path):
        log_file = tmp_path / "test.log"
        setup_logging(level="INFO", log_file=str(log_file))
        assert log_file.exists()

    def test_setup_creates_log_directory(self, tmp_path):
        log_file = tmp_path / "nested" / "subdir" / "test.log"
        setup_logging(level="INFO", log_file=str(log_file))
        assert log_file.parent.exists()
        assert log_file.exists()

    def test_setup_with_invalid_config_file(self):
        with patch("os.path.exists", return_value=True), patch(
            "builtins.open", side_effect=Exception("Bad config")
        ), patch("builtins.print") as mock_print:
            setup_logging(config_file="/bad_config.yaml")
            mock_print.assert_called_once()

    def test_setup_with_valid_config_file(self):
        config = {
            "version": 1,
            "disable_existing_loggers": False,
            "formatters": {
                "simple": {"format": "%(message)s"}
            },
            "handlers": {
                "console": {
                    "class": "logging.StreamHandler",
                    "level": "DEBUG",
                    "formatter": "simple",
                }
            },
            "root": {
                "level": "DEBUG",
                "handlers": ["console"],
            },
        }

        with patch("os.path.exists", return_value=True), patch(
            "builtins.open", return_value=MagicMock()
        ), patch("yaml.safe_load", return_value=config), patch(
            "logging.config.dictConfig"
        ) as mock_dict:
            setup_logging(config_file="/valid_config.yaml")
            mock_dict.assert_called_once_with(config)

    def test_respects_log_level(self):
        setup_logging(level="ERROR")
        root_logger = logging.getLogger()
        assert root_logger.level == logging.ERROR


class TestCreateDefaultLoggingConfig:
    def test_returns_dict(self):
        config = create_default_logging_config()
        assert isinstance(config, dict)

    def test_has_version_key(self):
        config = create_default_logging_config()
        assert config["version"] == 1

    def test_disables_existing_loggers(self):
        config = create_default_logging_config()
        assert config["disable_existing_loggers"] is False

    def test_has_formatters(self):
        config = create_default_logging_config()
        assert "formatters" in config
        assert "detailed" in config["formatters"]
        assert "simple" in config["formatters"]

    def test_has_handlers(self):
        config = create_default_logging_config()
        assert "handlers" in config
        assert "console" in config["handlers"]
        assert "file" in config["handlers"]

    def test_console_handler_config(self):
        config = create_default_logging_config()
        console = config["handlers"]["console"]
        assert console["class"] == "logging.StreamHandler"
        assert console["level"] == "INFO"

    def test_file_handler_config(self):
        config = create_default_logging_config()
        file_handler = config["handlers"]["file"]
        assert file_handler["class"] == "logging.handlers.RotatingFileHandler"
        assert file_handler["level"] == "DEBUG"
        assert file_handler["maxBytes"] == 10485760
        assert file_handler["backupCount"] == 5

    def test_has_loggers(self):
        config = create_default_logging_config()
        assert "loggers" in config
        assert "verificacion_correo" in config["loggers"]

    def test_verificacion_correo_logger_config(self):
        config = create_default_logging_config()
        logger_cfg = config["loggers"]["verificacion_correo"]
        assert logger_cfg["level"] == "DEBUG"
        assert "console" in logger_cfg["handlers"]
        assert "file" in logger_cfg["handlers"]
        assert logger_cfg["propagate"] is False

    def test_has_root_logger(self):
        config = create_default_logging_config()
        assert "root" in config
        assert config["root"]["level"] == "INFO"

    def test_detailed_format(self):
        config = create_default_logging_config()
        fmt = config["formatters"]["detailed"]["format"]
        assert "%(asctime)s" in fmt
        assert "%(name)s" in fmt
        assert "%(levelname)s" in fmt
        assert "%(message)s" in fmt


class TestIntegration:
    def test_get_logger_with_verificacion_correo_prefix(self):
        logger = get_logger("verificacion_correo.core.test")
        assert logger.name == "verificacion_correo.core.test"

    def test_module_logger_name_convention(self):
        logger = get_logger(__name__)
        assert logger.name == __name__

    def test_logger_levels(self):
        logger = get_logger("test_levels")
        logger.setLevel(logging.WARNING)
        assert logger.level == logging.WARNING
        assert logger.isEnabledFor(logging.ERROR)
        assert not logger.isEnabledFor(logging.DEBUG)
