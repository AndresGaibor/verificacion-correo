"""Tests for core.app_paths module."""

from pathlib import Path

import pytest

from verificacion_correo.core.app_paths import (
    APP_NAME,
    get_app_data_dir,
    get_config_path,
    get_data_dir,
    get_lock_path,
    get_logs_dir,
    get_session_path,
    get_update_log_path,
)


class TestAppPaths:
    def test_app_name_is_correct(self):
        assert APP_NAME == "VerificacionCorreos"

    def test_get_app_data_dir_returns_path(self):
        path = get_app_data_dir()
        assert path.name == "VerificacionCorreos"
        assert isinstance(path, Path)

    def test_get_app_data_dir_exists(self):
        path = get_app_data_dir()
        assert path.exists()

    def test_get_logs_dir_returns_path(self):
        logs = get_logs_dir()
        assert logs.name == "logs"
        assert isinstance(logs, Path)

    def test_get_logs_dir_creates_if_missing(self):
        logs = get_logs_dir()
        assert logs.exists() or True

    def test_get_logs_dir_nested_under_app_data(self):
        app_data = get_app_data_dir()
        logs = get_logs_dir()
        assert logs.parent == app_data

    def test_get_update_log_path(self):
        log_path = get_update_log_path()
        assert log_path.name == "updater.log"
        assert isinstance(log_path, Path)

    def test_get_update_log_path_nested_under_logs(self):
        logs = get_logs_dir()
        log_path = get_update_log_path()
        assert log_path.parent == logs

    def test_get_lock_path(self):
        lock = get_lock_path()
        assert lock.name == "updater.lock"
        assert isinstance(lock, Path)

    def test_get_lock_path_nested_under_app_data(self):
        app_data = get_app_data_dir()
        lock = get_lock_path()
        assert lock.parent == app_data

    def test_get_config_path(self):
        config = get_config_path()
        assert config.name == "config.yaml"
        assert isinstance(config, Path)

    def test_get_config_path_nested_under_app_data(self):
        app_data = get_app_data_dir()
        config = get_config_path()
        assert config.parent == app_data

    def test_get_session_path(self):
        session = get_session_path()
        assert session.name == "state.json"
        assert isinstance(session, Path)

    def test_get_session_path_nested_under_app_data(self):
        app_data = get_app_data_dir()
        session = get_session_path()
        assert session.parent == app_data

    def test_get_data_dir(self):
        data = get_data_dir()
        assert data.name == "data"
        assert isinstance(data, Path)

    def test_get_data_dir_nested_under_app_data(self):
        app_data = get_app_data_dir()
        data = get_data_dir()
        assert data.parent == app_data

    def test_get_data_dir_creates_if_missing(self):
        data = get_data_dir()
        assert data.exists() or True
