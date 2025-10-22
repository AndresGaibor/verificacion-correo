"""
Logging configuration and utilities for verificacion-correo.

This module provides centralized logging setup with configurable levels
and formatted output for both console and file logging.
"""

import logging
import logging.config
import os
from pathlib import Path
from typing import Optional, Dict, Any
import yaml


def setup_logging(
    level: str = "INFO",
    log_file: Optional[str] = None,
    config_file: Optional[str] = None
) -> None:
    """
    Setup logging configuration.

    Args:
        level: Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_file: Path to log file. If None, only console logging.
        config_file: Path to logging config YAML file.
    """
    # Try to load from config file first
    if config_file and os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
            logging.config.dictConfig(config)
            return
        except Exception as e:
            print(f"⚠️ Error loading logging config from {config_file}: {e}")

    # Fall back to basic configuration
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    date_format = '%Y-%m-%d %H:%M:%S'

    # Create formatters
    formatter = logging.Formatter(log_format, date_format)

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(getattr(logging, level.upper()))
    console_handler.setFormatter(formatter)

    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(getattr(logging, level.upper()))
    root_logger.addHandler(console_handler)

    # Add file handler if specified
    if log_file:
        # Ensure log directory exists
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)

        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)  # File gets all messages
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)


def get_logger(name: str) -> logging.Logger:
    """
    Get a logger instance with the specified name.

    Args:
        name: Logger name (usually __name__)

    Returns:
        Logger instance.
    """
    return logging.getLogger(name)


def create_default_logging_config() -> Dict[str, Any]:
    """
    Create a default logging configuration dictionary.

    Returns:
        Dictionary suitable for logging.config.dictConfig().
    """
    return {
        'version': 1,
        'disable_existing_loggers': False,
        'formatters': {
            'detailed': {
                'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                'datefmt': '%Y-%m-%d %H:%M:%S'
            },
            'simple': {
                'format': '%(levelname)s - %(message)s'
            },
        },
        'handlers': {
            'console': {
                'class': 'logging.StreamHandler',
                'level': 'INFO',
                'formatter': 'simple',
                'stream': 'ext://sys.stdout'
            },
            'file': {
                'class': 'logging.handlers.RotatingFileHandler',
                'level': 'DEBUG',
                'formatter': 'detailed',
                'filename': 'logs/verificacion-correo.log',
                'maxBytes': 10485760,  # 10MB
                'backupCount': 5,
                'encoding': 'utf8'
            }
        },
        'loggers': {
            'verificacion_correo': {
                'level': 'DEBUG',
                'handlers': ['console', 'file'],
                'propagate': False
            }
        },
        'root': {
            'level': 'INFO',
            'handlers': ['console']
        }
    }


# Initialize default logging on module import
setup_logging()