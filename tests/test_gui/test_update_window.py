"""Tests for gui.update_window module."""

from unittest.mock import MagicMock, patch
import pytest
import tkinter as tk

from verificacion_correo.gui.update_window import UpdateWindow
from verificacion_correo.core.update_models import UpdateStatus, UpdateResult


class TestUpdateWindowCreates:
    def test_update_window_creates(self):
        root = tk.Tk()
        window = UpdateWindow(root)
        assert window.status_var.get() == "Buscando actualizaciones…"
        root.destroy()

    def test_window_geometry(self):
        root = tk.Tk()
        window = UpdateWindow(root)
        assert window.root.geometry() == "400x180"
        root.destroy()

    def test_window_not_resizable(self):
        root = tk.Tk()
        window = UpdateWindow(root)
        assert window.root.resizable(False, False) is None
        root.destroy()


class TestUpdateWindowStartCheck:
    @patch('verificacion_correo.gui.update_window.check_for_updates')
    def test_start_check_initializes_thread(self, mock_check):
        mock_check.return_value = UpdateResult(
            status=UpdateStatus.SIN_ACTUALIZACION,
            message="No hay actualizaciones"
        )
        root = tk.Tk()
        window = UpdateWindow(root)
        callback = MagicMock()
        window.start_check(callback)
        window.root.update()
        window.close()
        mock_check.assert_called_once()


class TestUpdateWindowShowResult:
    def test_show_result_actualizado(self):
        root = tk.Tk()
        window = UpdateWindow(root)
        window.progress.grid_remove = MagicMock()
        result = UpdateResult(
            status=UpdateStatus.ACTUALIZADO,
            message="Actualización completada",
            previous_commit="abc123def456",
            current_commit="789xyz012345",
            commits_updated=5
        )
        window.show_result(result)
        assert window.status_var.get() == "Actualización instalada"
        root.destroy()

    def test_show_result_no_bloqueante(self):
        root = tk.Tk()
        window = UpdateWindow(root)
        window.progress.grid_remove = MagicMock()
        result = UpdateResult(
            status=UpdateStatus.SIN_ACTUALIZACION,
            message="No hay actualizaciones"
        )
        window.show_result(result)
        assert window.status_var.get() == "No hay actualizaciones"
        root.destroy()

    def test_show_result_error(self):
        root = tk.Tk()
        window = UpdateWindow(root)
        window.progress.grid_remove = MagicMock()
        result = UpdateResult(
            status=UpdateStatus.ERROR,
            message="Error al actualizar"
        )
        window.show_result(result)
        assert window.status_var.get() == "No fue posible actualizar"
        root.destroy()


class TestUpdateWindowClose:
    def test_close_cancels_after_id(self):
        root = tk.Tk()
        window = UpdateWindow(root)
        window._after_id = "some_id"
        window.root.after_cancel = MagicMock()
        window.close()
        window.root.after_cancel.assert_called_once_with("some_id")
        root.destroy()

    def test_close_without_after_id(self):
        root = tk.Tk()
        window = UpdateWindow(root)
        window._after_id = None
        window.root.after_cancel = MagicMock()
        window.close()
        window.root.after_cancel.assert_not_called()
