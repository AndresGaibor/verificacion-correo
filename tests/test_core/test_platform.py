"""Tests para verificacion_correo.core.platform.

Funciones cross-platform para abrir archivos y carpetas.
"""

import subprocess
import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock

from verificacion_correo.core.platform import open_in_explorer, open_file, open_folder


# ---------------------------------------------------------------------------
# open_in_explorer
# ---------------------------------------------------------------------------

class TestOpenInExplorer:
    """Tests para open_in_explorer()."""

    @patch("verificacion_correo.core.platform.platform")
    @patch("verificacion_correo.core.platform.subprocess")
    def test_darwin_calls_open(self, mock_subprocess, mock_platform, tmp_path):
        test_file = tmp_path / "test.txt"
        test_file.touch()
        mock_platform.system.return_value = "Darwin"
        mock_subprocess.run.return_value = MagicMock(returncode=0)

        result = open_in_explorer(test_file)

        assert result is True
        mock_subprocess.run.assert_called_once_with(
            ["open", str(test_file)],
            capture_output=True,
            timeout=10,
        )

    @patch("verificacion_correo.core.platform.platform")
    @patch("verificacion_correo.core.platform.subprocess")
    def test_linux_calls_xdg_open(self, mock_subprocess, mock_platform, tmp_path):
        test_file = tmp_path / "test.txt"
        test_file.touch()
        mock_platform.system.return_value = "Linux"
        mock_subprocess.run.return_value = MagicMock(returncode=0)

        result = open_in_explorer(test_file)

        assert result is True
        mock_subprocess.run.assert_called_once_with(
            ["xdg-open", str(test_file)],
            capture_output=True,
            timeout=10,
        )

    @patch("verificacion_correo.core.platform.platform")
    @patch("verificacion_correo.core.platform.os")
    def test_windows_calls_startfile(self, mock_os, mock_platform, tmp_path):
        test_file = tmp_path / "test.txt"
        test_file.touch()
        mock_platform.system.return_value = "Windows"

        result = open_in_explorer(test_file)

        assert result is True
        mock_os.startfile.assert_called_once_with(str(test_file))

    @patch("verificacion_correo.core.platform.platform")
    @patch("verificacion_correo.core.platform.os")
    def test_windows_startfile_error_returns_false(self, mock_os, mock_platform, tmp_path):
        test_file = tmp_path / "test.txt"
        test_file.touch()
        mock_platform.system.return_value = "Windows"
        mock_os.startfile.side_effect = OSError("File not found")

        result = open_in_explorer(test_file)

        assert result is False

    @patch("verificacion_correo.core.platform.platform")
    @patch("verificacion_correo.core.platform.subprocess")
    def test_subprocess_file_not_found_returns_false(self, mock_subprocess, mock_platform, tmp_path):
        test_file = tmp_path / "test.txt"
        test_file.touch()
        mock_platform.system.return_value = "Darwin"
        mock_subprocess.run.side_effect = FileNotFoundError("open not found")

        result = open_in_explorer(test_file)

        assert result is False

    @patch("verificacion_correo.core.platform.platform")
    @patch("verificacion_correo.core.platform.subprocess")
    def test_subprocess_timeout_returns_true(self, mock_subprocess, mock_platform, tmp_path):
        test_file = tmp_path / "test.txt"
        test_file.touch()
        mock_platform.system.return_value = "Darwin"
        mock_subprocess.run.side_effect = subprocess.TimeoutExpired(cmd="open", timeout=10)

        result = open_in_explorer(test_file)

        # Timeout means the command was launched — considered success
        assert result is True

    @patch("verificacion_correo.core.platform.platform")
    @patch("verificacion_correo.core.platform.subprocess")
    def test_nonzero_return_returns_false(self, mock_subprocess, mock_platform, tmp_path):
        test_file = tmp_path / "test.txt"
        test_file.touch()
        mock_platform.system.return_value = "Linux"
        mock_subprocess.run.return_value = MagicMock(returncode=1)

        result = open_in_explorer(test_file)

        assert result is False

    def test_returns_false_for_nonexistent_file(self, tmp_path):
        """Si el path parece un archivo (con extensión) y no existe, retorna False."""
        result = open_in_explorer(tmp_path / "nonexistent.txt")
        assert result is False


# ---------------------------------------------------------------------------
# open_file
# ---------------------------------------------------------------------------

class TestOpenFile:
    """Tests para open_file()."""

    @patch("verificacion_correo.core.platform.open_in_explorer")
    def test_returns_true_when_open_succeeds(self, mock_open, tmp_path):
        test_file = tmp_path / "test.xlsx"
        test_file.touch()
        mock_open.return_value = True

        result = open_file(test_file)

        assert result is True
        mock_open.assert_called_once_with(test_file)

    @patch("verificacion_correo.core.platform.open_in_explorer")
    def test_returns_false_when_open_fails(self, mock_open, tmp_path):
        test_file = tmp_path / "test.xlsx"
        test_file.touch()
        mock_open.return_value = False

        result = open_file(test_file)

        assert result is False

    def test_returns_false_for_nonexistent_file(self, tmp_path):
        result = open_file(tmp_path / "no_existe.xlsx")
        assert result is False

    def test_returns_false_for_directory(self, tmp_path):
        result = open_file(tmp_path)
        assert result is False

    def test_accepts_string_path(self, tmp_path):
        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        with patch("verificacion_correo.core.platform.open_in_explorer", return_value=True) as mock_open:
            result = open_file(str(test_file))
            assert result is True


# ---------------------------------------------------------------------------
# open_folder
# ---------------------------------------------------------------------------

class TestOpenFolder:
    """Tests para open_folder()."""

    @patch("verificacion_correo.core.platform.open_in_explorer")
    def test_returns_true_when_open_succeeds(self, mock_open, tmp_path):
        mock_open.return_value = True

        result = open_folder(tmp_path)

        assert result is True
        mock_open.assert_called_once_with(tmp_path)

    @patch("verificacion_correo.core.platform.open_in_explorer")
    def test_returns_false_when_open_fails(self, mock_open, tmp_path):
        mock_open.return_value = False

        result = open_folder(tmp_path)

        assert result is False

    def test_returns_false_for_file_path(self, tmp_path):
        test_file = tmp_path / "file.txt"
        test_file.touch()

        result = open_folder(test_file)
        assert result is False

    def test_accepts_string_path(self, tmp_path):
        with patch("verificacion_correo.core.platform.open_in_explorer", return_value=True) as mock_open:
            result = open_folder(str(tmp_path))
            assert result is True


# ---------------------------------------------------------------------------
# Import check
# ---------------------------------------------------------------------------

class TestPlatformImport:
    """Verificar que el módulo se puede importar correctamente."""

    def test_import_from_core(self):
        from verificacion_correo.core.platform import open_in_explorer, open_file, open_folder
        assert callable(open_in_explorer)
        assert callable(open_file)
        assert callable(open_folder)

    def test_import_from_core_init(self):
        from verificacion_correo.core import open_file, open_folder, open_in_explorer
        assert callable(open_file)
        assert callable(open_folder)
        assert callable(open_in_explorer)
