"""Tests for core.session module."""

import json
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

from verificacion_correo.core.session import SessionManager, get_session_status


SESSION_DATA = {
    "cookies": [
        {"name": "X-OWA-CANARY", "value": "test_canary_123", "domain": ".madrid.org"},
        {"name": "session_id", "value": "abc123", "domain": ".madrid.org"},
    ],
    "origins": [
        {
            "origin": "https://correoweb.madrid.org",
            "localStorage": [{"canary_key": "local_canary"}],
        }
    ],
}


def _make_config(session_path: str) -> MagicMock:
    """Create a mock Config with the given session file path."""
    config = MagicMock()
    config.get_session_file_path.return_value = session_path
    return config


class TestSessionManagerInit:
    def test_session_file_resolved_to_absolute_path(self, tmp_path):
        session_file = tmp_path / "state.json"
        config = _make_config(str(session_file))
        manager = SessionManager(config)
        assert manager.session_file == session_file
        assert manager.session_file.is_absolute()

    def test_session_file_preserves_config_path(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        config = _make_config("state.json")
        manager = SessionManager(config)
        assert manager.session_file == Path("state.json")
        assert manager.session_file.name == "state.json"


class TestDeleteSession:
    def test_delete_session_removes_file(self, tmp_path):
        session_file = tmp_path / "state.json"
        session_file.write_text("{}")
        config = _make_config(str(session_file))
        manager = SessionManager(config)

        result = manager.delete_session()

        assert result is True
        assert not session_file.exists()

    def test_delete_session_returns_true_when_no_file(self, tmp_path):
        session_file = tmp_path / "nonexistent.json"
        config = _make_config(str(session_file))
        manager = SessionManager(config)

        result = manager.delete_session()

        assert result is True


class TestGetSessionStatus:
    def test_session_not_found(self, tmp_path):
        session_file = tmp_path / "nonexistent.json"
        config = _make_config(str(session_file))

        status = get_session_status(config)

        assert status["exists"] is False
        assert status["is_valid"] is False

    def test_session_expired_empty_file(self, tmp_path):
        session_file = tmp_path / "state.json"
        session_file.write_text("")
        config = _make_config(str(session_file))

        status = get_session_status(config)

        assert status["exists"] is True
        assert "error" in status
        assert status["is_valid"] is False

    def test_session_valid_structure(self, tmp_path):
        session_file = tmp_path / "state.json"
        session_file.write_text(json.dumps(SESSION_DATA))
        config = _make_config(str(session_file))

        with patch.object(
            SessionManager, "validate_session", return_value=True
        ):
            status = get_session_status(config)

        assert status["exists"] is True
        assert status["cookies_count"] == 2
        assert status["origins_count"] == 1
        assert status["is_valid"] is True

    def test_session_no_cookies(self, tmp_path):
        session_file = tmp_path / "state.json"
        session_file.write_text(json.dumps({"cookies": [], "origins": []}))
        config = _make_config(str(session_file))

        with patch.object(
            SessionManager, "validate_session", return_value=False
        ):
            status = get_session_status(config)

        assert status["exists"] is True
        assert status["cookies_count"] == 0
        assert status["origins_count"] == 0
