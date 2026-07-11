"""Tests for core.first_run module."""

from pathlib import Path
from unittest.mock import MagicMock

import pytest

from verificacion_correo.core.first_run import FirstRunManager


@pytest.fixture
def manager():
    """Create a fresh FirstRunManager."""
    return FirstRunManager()


class TestFirstRunManagerInit:
    def test_init_sets_config_to_none(self, manager):
        assert manager.config is None


class TestGetMarkerPath:
    def test_marker_path_without_config(self, manager):
        marker = manager._get_marker_path()
        assert marker == Path(".first_run_completed")

    def test_marker_path_with_config(self, manager):
        mock_config = MagicMock()
        mock_config.get_excel_file_path.return_value = "/tmp/data/correos.xlsx"
        manager.config = mock_config

        marker = manager._get_marker_path()
        assert marker == Path("/tmp/data/.first_run_completed")
        assert marker.name == ".first_run_completed"


class TestIsFirstRun:
    def test_is_first_run_true_when_no_config(self, manager, tmp_path, monkeypatch):
        """Returns True when no config files nor marker exist."""
        monkeypatch.chdir(tmp_path)
        assert manager.is_first_run() is True

    def test_is_first_run_false_when_config_yaml_exists(self, manager, tmp_path, monkeypatch):
        """Returns False when config.yaml exists."""
        monkeypatch.chdir(tmp_path)
        (tmp_path / "config.yaml").touch()
        assert manager.is_first_run() is False

    def test_is_first_run_false_when_default_yaml_exists(self, manager, tmp_path, monkeypatch):
        """Returns False when config/default.yaml exists."""
        monkeypatch.chdir(tmp_path)
        config_dir = tmp_path / "config"
        config_dir.mkdir()
        (config_dir / "default.yaml").touch()
        assert manager.is_first_run() is False

    def test_is_first_run_false_when_marker_exists(self, manager, tmp_path, monkeypatch):
        """Returns False when marker file exists."""
        monkeypatch.chdir(tmp_path)
        (tmp_path / ".first_run_completed").touch()
        assert manager.is_first_run() is False

    def test_is_first_run_false_when_config_yaml_checked_first(self, manager, tmp_path, monkeypatch):
        """config.yaml takes precedence — even if marker would also return False."""
        monkeypatch.chdir(tmp_path)
        (tmp_path / "config.yaml").touch()
        (tmp_path / ".first_run_completed").touch()
        assert manager.is_first_run() is False


class TestEnsureDirectories:
    def test_ensure_directories_creates_dirs(self, manager, tmp_path, monkeypatch):
        """Creates data/, logs/, sessions/, exports/ directories."""
        monkeypatch.chdir(tmp_path)

        manager._ensure_directories()

        assert (tmp_path / "data").is_dir()
        assert (tmp_path / "logs").is_dir()
        assert (tmp_path / "sessions").is_dir()
        assert (tmp_path / "exports").is_dir()

    def test_ensure_directories_idempotent(self, manager, tmp_path, monkeypatch):
        """Calling twice does not fail."""
        monkeypatch.chdir(tmp_path)

        manager._ensure_directories()
        manager._ensure_directories()

        assert (tmp_path / "data").is_dir()


class TestCreateFirstRunMarker:
    def test_create_first_run_marker(self, manager, tmp_path, monkeypatch):
        """Creates marker file with config info."""
        monkeypatch.chdir(tmp_path)

        mock_config = MagicMock()
        mock_config._config_path = str(tmp_path / "config.yaml")
        mock_config.get_excel_file_path.return_value = str(tmp_path / "data" / "correos.xlsx")
        mock_config.get_session_file_path.return_value = str(tmp_path / "session.json")
        manager.config = mock_config

        manager._create_first_run_marker()

        marker = tmp_path / "data" / ".first_run_completed"
        assert marker.exists()
        content = marker.read_text()
        assert "First run completed" in content
        assert "config.yaml" in content

    def test_create_first_run_marker_with_config_creates_in_data_dir(self, manager, tmp_path, monkeypatch):
        """When config exists, marker is created in data/ dir."""
        monkeypatch.chdir(tmp_path)

        data_dir = tmp_path / "data"
        data_dir.mkdir()

        mock_config = MagicMock()
        mock_config._config_path = str(tmp_path / "config.yaml")
        mock_config.get_excel_file_path.return_value = str(data_dir / "correos.xlsx")
        mock_config.get_session_file_path.return_value = str(tmp_path / "session.json")
        manager.config = mock_config

        manager._create_first_run_marker()

        marker = data_dir / ".first_run_completed"
        assert marker.exists()

    def test_create_first_run_marker_with_relative_paths(self, manager, tmp_path, monkeypatch):
        """When config is set with relative path, marker is created in data/ dir."""
        monkeypatch.chdir(tmp_path)

        data_dir = tmp_path / "data"
        data_dir.mkdir()

        mock_config = MagicMock()
        mock_config._config_path = "config.yaml"
        mock_config.get_excel_file_path.return_value = "data/correos.xlsx"
        mock_config.get_session_file_path.return_value = "session.json"
        manager.config = mock_config

        manager._create_first_run_marker()

        marker = data_dir / ".first_run_completed"
        assert marker.exists()


class TestShowFirstRunSummary:
    def test_show_first_run_summary_no_config(self, manager, capsys):
        """Does nothing when config is None."""
        manager.show_first_run_summary()
        captured = capsys.readouterr()
        assert captured.out == ""

    def test_show_first_run_summary_with_config(self, manager, capsys):
        """Prints summary when config is set."""
        mock_config = MagicMock()
        mock_config._config_path = "config.yaml"
        mock_config.page_url = "https://example.com/owa"
        mock_config.get_excel_file_path.return_value = "data/correos.xlsx"
        mock_config.get_session_file_path.return_value = "session.json"
        mock_config.processing.batch_size = 10
        mock_config.default_emails = ["test@example.com"]
        manager.config = mock_config

        manager.show_first_run_summary()
        captured = capsys.readouterr()
        assert "config.yaml" in captured.out
        assert "https://example.com/owa" in captured.out
