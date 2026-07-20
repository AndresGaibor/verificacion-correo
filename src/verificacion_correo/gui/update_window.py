import tkinter as tk
from tkinter import ttk
from typing import Callable, Optional
from pathlib import Path
import threading
import queue

from ..core.update_models import UpdateStatus, UpdateResult, ERRORES_NO_BLOQUEANTES
from ..core.updater import check_for_updates, apply_update

class UpdateWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Verificación de Correos")
        self.root.geometry("400x180")
        self.root.resizable(False, False)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.frame = ttk.Frame(root, padding="20")
        self.frame.grid(column=0, row=0, sticky="NSEW")
        self.frame.columnconfigure(0, weight=1)

        self.title_label = ttk.Label(self.frame, text="Verificación de Correos", font=("Segoe UI", 14, "bold"))
        self.title_label.grid(column=0, row=0, pady=(0, 10))

        self.status_var = tk.StringVar(value="Buscando actualizaciones…")
        self.status_label = ttk.Label(self.frame, textvariable=self.status_var, font=("Segoe UI", 10))
        self.status_label.grid(column=0, row=1, pady=(0, 15))

        self.progress = ttk.Progressbar(self.frame, mode="indeterminate", length=300)
        self.progress.grid(column=0, row=2, pady=(0, 10))
        self.progress.start(10)

        self.detail_var = tk.StringVar(value="")
        self.detail_label = ttk.Label(self.frame, textvariable=self.detail_var, font=("Segoe UI", 8), foreground="gray")
        self.detail_label.grid(column=0, row=3)

        self.queue: queue.Queue = queue.Queue()
        self._after_id = None

    def start_check(self, on_complete: Callable[[UpdateResult], None]):
        self.on_complete = on_complete
        self._thread = threading.Thread(target=self._check_thread, daemon=True)
        self._thread.start()
        self._check_queue()

    def _check_thread(self):
        result = check_for_updates()
        self.queue.put(result)

    def _check_queue(self):
        try:
            result = self.queue.get_nowait()
            self._on_check_complete(result)
            return
        except queue.Empty:
            self._after_id = self.root.after(100, self._check_queue)

    def _on_check_complete(self, result: UpdateResult):
        self.progress.stop()
        self.progress.grid_remove()

        if result.status == UpdateStatus.SIN_ACTUALIZACION and result.commits_updated > 0:
            self.status_var.set("Descargando cambios…")
            self.progress.grid()
            self.progress.start(10)
            self._apply_update_thread()
        elif result.status in ERRORES_NO_BLOQUEANTES:
            self.root.after(500, lambda: self.on_complete(result))
        else:
            self.root.after(500, lambda: self.on_complete(result))

    def _apply_update_thread(self):
        self._thread = threading.Thread(target=self._apply_thread, daemon=True)
        self._thread.start()

    def _apply_thread(self):
        result = apply_update()
        self.queue.put(result)
        self._check_apply_queue()

    def _check_apply_queue(self):
        try:
            result = self.queue.get_nowait()
            self._on_apply_complete(result)
            return
        except queue.Empty:
            self._after_id = self.root.after(100, self._check_apply_queue)

    def _on_apply_complete(self, result: UpdateResult):
        self.progress.stop()
        self.progress.grid_remove()
        self.root.after(500, lambda: self.on_complete(result))

    def show_result(self, result: UpdateResult):
        self.progress.grid_remove()

        if result.status == UpdateStatus.ACTUALIZADO:
            self.status_var.set("Actualización instalada")
            self.detail_var.set(
                f"Se actualizaron {result.commits_updated} cambios correctamente.\n"
                f"Versión anterior: {result.previous_commit[:7]}\n"
                f"Nueva versión: {result.current_commit[:7]}\n\n"
                f"La aplicación se cerrará en 3 segundos."
            )
            self.root.after(3000, self.root.destroy)
        elif result.status in ERRORES_NO_BLOQUEANTES:
            self.status_var.set("No hay actualizaciones")
            self.detail_var.set(result.message)
        else:
            self.status_var.set("No fue posible actualizar")
            self.detail_var.set(result.message)

    def close(self):
        if self._after_id:
            self.root.after_cancel(self._after_id)
        self.root.destroy()
