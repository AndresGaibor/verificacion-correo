import os
import platform
from pathlib import Path

APP_NAME = "VerificacionCorreos"


def _get_base_dir() -> Path:
    if platform.system() == "Windows":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    else:
        base = Path.home() / ".config"
    return base / APP_NAME


def get_app_data_dir() -> Path:
    d = _get_base_dir()
    d.mkdir(parents=True, exist_ok=True)
    return d


def get_logs_dir() -> Path:
    d = get_app_data_dir() / "logs"
    d.mkdir(parents=True, exist_ok=True)
    return d


def get_update_log_path() -> Path:
    return get_logs_dir() / "updater.log"


def get_lock_path() -> Path:
    return get_app_data_dir() / "updater.lock"


def get_config_path() -> Path:
    return get_app_data_dir() / "config.yaml"


def get_session_path() -> Path:
    return get_app_data_dir() / "state.json"


def get_data_dir() -> Path:
    d = get_app_data_dir() / "data"
    d.mkdir(parents=True, exist_ok=True)
    return d
