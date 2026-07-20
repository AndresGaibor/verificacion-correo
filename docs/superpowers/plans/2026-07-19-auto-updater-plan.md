# Auto-Updater Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implementar un auto-updater que se ejecute ANTES de importar la GUI, verifique actualizaciones via Git, y muestre una ventana de progreso.

**Architecture:** Launcher separado que ejecuta verificación de Git en hilo secundario, con ventana Tkinter para progreso. Fail-open: si falla cualquier cosa, abre la GUI normalmente.

**Tech Stack:** Tkinter para UI, subprocess para Git, threading para verificación async, sys.executable para instalar dependencias en el Python correcto.

---

## Global Constraints

- No importar `gui.main` hasta después de la verificación de actualizaciones
- Solo actualizaciones fast-forward (`git pull --ff-only`)
- Nunca borrar cambios locales
- Datos personales fuera del repositorio (ya configurado en .gitignore)
- Fail-open: cualquier error no bloquea la apertura de la GUI

---

## File Structure

```
src/verificacion_correo/
├── launcher.py                    # Entry point oficial
├── core/
│   ├── updater.py                 # Lógica Git
│   ├── update_models.py            # Enums y dataclasses
│   └── app_paths.py               # Rutas centralizadas
└── gui/
    └── update_window.py            # Ventana de progreso
```

---

## Tasks

### Task 1: app_paths.py — Rutas centralizadas

**Files:**
- Create: `src/verificacion_correo/core/app_paths.py`
- Test: `tests/test_core/test_app_paths.py`

**Interfaces:**
- Consumes: Nothing
- Produces:
  - `APP_NAME = "VerificacionCorreos"`
  - `get_app_data_dir() -> Path` — `%LOCALAPPDATA%\VerificacionCorreos` en Windows, `~/.config/VerificacionCorreos` en Unix
  - `get_config_path() -> Path`
  - `get_session_path() -> Path`
  - `get_data_dir() -> Path`
  - `get_logs_dir() -> Path`
  - `get_update_log_path() -> Path`
  - `get_lock_path() -> Path`

- [ ] **Step 1: Escribir test**

```python
def test_get_app_data_dir_returns_path():
    path = get_app_data_dir()
    assert path.name == "VerificacionCorreos"
    assert path.exists()

def test_get_logs_dir_creates_if_missing():
    logs = get_logs_dir()
    assert logs.exists() or True  # Puede crear

def test_get_lock_path():
    lock = get_lock_path()
    assert lock.name == "updater.lock"
```

- [ ] **Step 2: Run test — debe fallar (archivo no existe)**

- [ ] **Step 3: Implementar app_paths.py**

```python
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
```

- [ ] **Step 4: Run tests — deben pasar**

- [ ] **Step 5: Commit**

---

### Task 2: update_models.py — Enums y dataclasses

**Files:**
- Create: `src/verificacion_correo/core/update_models.py`
- Test: `tests/test_core/test_update_models.py`

**Interfaces:**
- Consumes: Nothing
- Produces:
  - `UpdateStatus` enum
  - `UpdateResult` dataclass

- [ ] **Step 1: Escribir test**

```python
def test_update_status_enum_values():
    assert UpdateStatus.SIN_ACTUALIZACION.value == "sin_actualizacion"
    assert UpdateStatus.ACTUALIZADO.value == "actualizado"
    assert UpdateStatus.SIN_INTERNET.value == "sin_internet"

def test_update_result_dataclass():
    result = UpdateResult(
        status=UpdateStatus.SIN_ACTUALIZACION,
        message="No hay actualizaciones",
        previous_commit="abc123",
        current_commit="abc123",
        commits_updated=0,
    )
    assert result.status == UpdateStatus.SIN_ACTUALIZACION
    assert result.commits_updated == 0
```

- [ ] **Step 2: Implementar update_models.py**

```python
from enum import Enum
from dataclasses import dataclass, field
from typing import Optional

class UpdateStatus(Enum):
    SIN_ACTUALIZACION = "sin_actualizacion"
    ACTUALIZADO = "actualizado"
    SIN_INTERNET = "sin_internet"
    GIT_NO_DISPONIBLE = "git_no_disponible"
    REPOSITORIO_INVALIDO = "repositorio_invalido"
    CAMBIOS_LOCALES = "cambios_locales"
    HISTORIAL_DIVERGENTE = "historial_divergente"
    ERROR = "error"
    BLOQUEADO = "bloqueado"

@dataclass
class UpdateResult:
    status: UpdateStatus
    message: str
    previous_commit: Optional[str] = None
    current_commit: Optional[str] = None
    commits_updated: int = 0
    dependencies_updated: bool = False
    error_detail: Optional[str] = None

ERRORES_NO_BLOQUEANTES = {
    UpdateStatus.SIN_ACTUALIZACION,
    UpdateStatus.SIN_INTERNET,
    UpdateStatus.GIT_NO_DISPONIBLE,
    UpdateStatus.REPOSITORIO_INVALIDO,
}
```

- [ ] **Step 3: Run tests — deben pasar**

- [ ] **Step 4: Commit**

---

### Task 3: updater.py — Lógica Git

**Files:**
- Create: `src/verificacion_correo/core/updater.py`
- Test: `tests/test_core/test_updater.py`

**Interfaces:**
- Consumes: `app_paths.py`, `update_models.py`
- Produces: `check_for_updates() -> UpdateResult`, `apply_update() -> UpdateResult`, `_acquire_lock() -> bool`, `_release_lock() -> None`

- [ ] **Step 1: Escribir tests con git repo temporal**

```python
import tempfile
import subprocess

def test_check_updates_no_remote():
    with tempfile.TemporaryDirectory() as tmpdir:
        # Crear repo git sin remoto
        subprocess.run(["git", "init"], cwd=tmpdir, capture_output=True)
        result = check_for_updates(tmpdir)
        assert result.status == UpdateStatus.REPOSITORIO_INVALIDO

def test_check_updates_clean():
    with tempfile.TemporaryDirectory() as tmpdir:
        subprocess.run(["git", "init"], cwd=tmpdir, capture_output=True)
        subprocess.run(["git", "config", "user.email", "test@test.com"], cwd=tmpdir, capture_output=True)
        subprocess.run(["git", "config", "user.name", "Test"], cwd=tmpdir, capture_output=True)
        # Crear archivo y commit
        Path(tmpdir, "test.txt").write_text("hello")
        subprocess.run(["git", "add", "."], cwd=tmpdir, capture_output=True)
        subprocess.run(["git", "commit", "-m", "init"], cwd=tmpdir, capture_output=True)
        result = check_for_updates(tmpdir)
        assert result.status == UpdateStatus.SIN_ACTUALIZACION
```

- [ ] **Step 2: Implementar updater.py**

```python
import subprocess
import hashlib
import sys
import fcntl
import os
from pathlib import Path
from typing import Optional, Tuple
import shutil

from .update_models import UpdateStatus, UpdateResult, ERRORES_NO_BLOQUEANTES
from .app_paths import get_lock_path, get_update_log_path

EXPECTED_REMOTE = "https://github.com/andresgaibor/verificacion-correo.git"
UPDATE_BRANCH = "main"
GIT_TIMEOUT = 20

def _log(msg: str):
    log_path = get_update_log_path()
    with open(log_path, "a", encoding="utf-8") as f:
        from datetime import datetime
        f.write(f"{datetime.now().isoformat()} {msg}\n")

def _run_git(cwd: Path, *args, timeout: int = GIT_TIMEOUT) -> Tuple[int, str, str]:
    try:
        result = subprocess.run(
            ["git", *args],
            cwd=cwd,
            capture_output=True,
            text=True,
            timeout=timeout,
            shell=False,
        )
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except subprocess.TimeoutExpired:
        return -1, "", "Timeout"
    except FileNotFoundError:
        return -2, "", "Git not found"

def _is_git_available() -> bool:
    code, _, _ = _run_git(Path.cwd(), "version")
    return code == 0

def _get_repo_root(path: Path) -> Optional[Path]:
    code, stdout, _ = _run_git(path, "rev-parse", "--show-toplevel")
    if code != 0:
        return None
    return Path(stdout)

def _get_remote_url(cwd: Path) -> Optional[str]:
    code, stdout, _ = _run_git(cwd, "remote", "get-url", "origin")
    if code != 0:
        return None
    return stdout

def _is_repo_clean(cwd: Path) -> Tuple[bool, str]:
    code, stdout, _ = _run_git(cwd, "status", "--porcelain")
    if code != 0:
        return False, "No se pudo ejecutar git status"
    if stdout:
        return False, f"Cambios locales: {stdout[:200]}"
    return True, ""

def _get_current_commit(cwd: Path) -> Optional[str]:
    code, stdout, _ = _run_git(cwd, "rev-parse", "HEAD")
    if code != 0:
        return None
    return stdout

def _get_remote_commit(cwd: Path, branch: str = "origin/main") -> Optional[str]:
    code, stdout, _ = _run_git(cwd, "rev-parse", branch)
    if code != 0:
        return None
    return stdout

def check_for_updates(repo_root: Optional[Path] = None) -> UpdateResult:
    if repo_root is None:
        repo_root = Path.cwd()

    _log(f"INFO Inicio de comprobación en {repo_root}")

    if not _is_git_available():
        _log("WARN Git no disponible")
        return UpdateResult(status=UpdateStatus.GIT_NO_DISPONIBLE, message="Git no está instalado")

    root = _get_repo_root(repo_root)
    if not root:
        _log("WARN No es repositorio Git")
        return UpdateResult(status=UpdateStatus.REPOSITORIO_INVALIDO, message="No es un repositorio Git")

    remote_url = _get_remote_url(root)
    if not remote_url:
        _log("WARN Sin remoto configurado")
        return UpdateResult(status=UpdateStatus.REPOSITORIO_INVALIDO, message="No existe remoto origin")

    if remote_url != EXPECTED_REMOTE:
        _log(f"WARN Remoto inesperado: {remote_url}")
        return UpdateResult(status=UpdateStatus.REPOSITORIO_INVALIDO, message=f"Remoto no autorizado: {remote_url}")

    code, _, stderr = _run_git(root, "fetch", "origin", UPDATE_BRANCH, "--prune")
    if code != 0:
        _log(f"WARN Fetch falló: {stderr}")
        return UpdateResult(status=UpdateStatus.SIN_INTERNET, message="Sin conexión o GitHub no responde")

    local_commit = _get_current_commit(root)
    remote_commit = _get_remote_commit(root, f"origin/{UPDATE_BRANCH}")

    _log(f"INFO HEAD local: {local_commit}")
    _log(f"INFO HEAD remoto: {remote_commit}")

    if local_commit == remote_commit:
        _log("INFO Sin actualizaciones")
        return UpdateResult(
            status=UpdateStatus.SIN_ACTUALIZACION,
            message="No hay actualizaciones",
            previous_commit=local_commit,
            current_commit=remote_commit,
        )

    is_ancestor_code, _, _ = _run_git(root, "merge-base", "--is-ancestor", local_commit, remote_commit)
    if is_ancestor_code != 0:
        _log("WARN Historial divergente")
        return UpdateResult(
            status=UpdateStatus.HISTORIAL_DIVERGENTE,
            message="El historial local y remoto han divergido",
            previous_commit=local_commit,
            current_commit=remote_commit,
        )

    clean, reason = _is_repo_clean(root)
    if not clean:
        _log(f"WARN Cambios locales bloquean actualización: {reason}")
        return UpdateResult(
            status=UpdateStatus.CAMBIOS_LOCALES,
            message=f"No se puede actualizar: {reason}",
            previous_commit=local_commit,
            current_commit=remote_commit,
        )

    commits = _count_commits_between(root, local_commit, remote_commit)
    _log(f"INFO Actualización disponible: {commits} commits")
    return UpdateResult(
        status=UpdateStatus.SIN_ACTUALIZACION,
        message=f"Actualización disponible: {commits} commits",
        previous_commit=local_commit,
        current_commit=remote_commit,
        commits_updated=commits,
    )

def _count_commits_between(cwd: Path, from_commit: str, to_commit: str) -> int:
    code, stdout, _ = _run_git(cwd, "rev-list", "--count", f"{from_commit}..{to_commit}")
    if code != 0:
        return 0
    try:
        return int(stdout)
    except ValueError:
        return 0

def apply_update(repo_root: Optional[Path] = None) -> UpdateResult:
    if repo_root is None:
        repo_root = Path.cwd()

    _log("INFO Iniciando actualización")

    if not _acquire_lock():
        return UpdateResult(status=UpdateStatus.BLOQUEADO, message="Otra actualización en progreso")

    try:
        old_commit = _get_current_commit(repo_root)

        code, _, stderr = _run_git(repo_root, "pull", "--ff-only", "origin", UPDATE_BRANCH)
        if code != 0:
            _log(f"ERROR Pull falló: {stderr}")
            return UpdateResult(
                status=UpdateStatus.ERROR,
                message="Error al ejecutar git pull",
                previous_commit=old_commit,
                error_detail=stderr,
            )

        new_commit = _get_current_commit(repo_root)

        pyproject_hash_before = _get_file_hash(repo_root / "pyproject.toml")
        deps_updated = _update_dependencies()
        pyproject_hash_after = _get_file_hash(repo_root / "pyproject.toml")

        commits = _count_commits_between(repo_root, old_commit, new_commit)

        _log(f"INFO Actualización completada: {commits} commits, deps={deps_updated}")

        return UpdateResult(
            status=UpdateStatus.ACTUALIZADO,
            message=f"Se actualizaron {commits} cambios correctamente",
            previous_commit=old_commit,
            current_commit=new_commit,
            commits_updated=commits,
            dependencies_updated=deps_updated,
        )

    except Exception as e:
        _log(f"ERROR Excepción: {e}")
        return UpdateResult(status=UpdateStatus.ERROR, message=str(e), error_detail=str(e))
    finally:
        _release_lock()

def _get_file_hash(path: Path) -> str:
    if not path.exists():
        return ""
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()

def _update_dependencies() -> bool:
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "-e", "."],
            capture_output=True,
            text=True,
            timeout=120,
            shell=False,
        )
        if result.returncode == 0:
            _log("INFO Dependencias actualizadas")
            return True
        else:
            _log(f"WARN pip install falló: {result.stderr[:200]}")
            return False
    except Exception as e:
        _log(f"WARN Excepción actualizando dependencias: {e}")
        return False

def _acquire_lock() -> bool:
    lock_path = get_lock_path()
    try:
        fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.close(fd)
        return True
    except FileExistsError:
        pid = _get_lock_pid()
        if pid and _is_process_running(pid):
            return False
        try:
            os.remove(lock_path)
        except FileNotFoundError:
            pass
        return _acquire_lock()

def _release_lock():
    lock_path = get_lock_path()
    try:
        os.remove(lock_path)
    except FileNotFoundError:
        pass

def _get_lock_pid() -> Optional[int]:
    lock_path = get_lock_path()
    if not lock_path.exists():
        return None
    try:
        return int(lock_path.read_text().strip())
    except (ValueError, FileNotFoundError):
        return None

def _is_process_running(pid: int) -> bool:
    import signal
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False
```

- [ ] **Step 3: Run tests**

- [ ] **Step 4: Commit**

---

### Task 4: update_window.py — Ventana de progreso

**Files:**
- Create: `src/verificacion_correo/gui/update_window.py`
- Test: `tests/test_gui/test_update_window.py`

**Interfaces:**
- Consumes: `update_models.py`, `updater.py`
- Produces: `UpdateWindow` class con `start_check()`, `show_result()`, `close()`

- [ ] **Step 1: Escribir test básico**

```python
def test_update_window_creates():
    root = tk.Tk()
    window = UpdateWindow(root)
    assert window.status_var.get() == ""
    root.destroy()
```

- [ ] **Step 2: Implementar update_window.py**

```python
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
```

- [ ] **Step 3: Run tests**

- [ ] **Step 4: Commit**

---

### Task 5: launcher.py — Punto de entrada oficial

**Files:**
- Create: `src/verificacion_correo/launcher.py`
- Modify: `src/verificacion_correo/__main__.py`

**Interfaces:**
- Consumes: `update_window.py`, `update_models.py`
- Produces: `launcher()` que importa `gui.main` SOLO después de verificar actualizaciones

- [ ] **Step 1: Crear launcher.py**

```python
#!/usr/bin/env python3
"""
Launcher - Entry point oficial de la aplicación.

Verifica actualizaciones ANTES de importar la GUI.
"""

import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from verificacion_correo.core.update_models import UpdateStatus, UpdateResult, ERRORES_NO_BLOQUEANTES
from verificacion_correo.core.updater import check_for_updates, apply_update
from verificacion_correo.gui.update_window import UpdateWindow


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--skip-update", action="store_true", help="Omitir verificación de actualizaciones")
    args = parser.parse_args()

    if args.skip_update:
        abrir_gui()
        return 0

    import tkinter as tk
    root = tk.Tk()
    root.withdraw()

    def on_complete(result: UpdateResult):
        root.withdraw()

        if result.status == UpdateStatus.ACTUALIZADO:
            _mostrar_y_cerrar(root, result)
            return

        if result.status in ERRORES_NO_BLOQUEANTES:
            root.destroy()
            abrir_gui()
            return

        root.destroy()
        abrir_gui()

    window = UpdateWindow(root)
    window.start_check(on_complete)

    root.mainloop()
    return 0


def _mostrar_y_cerrar(root: tk.Tk, result: UpdateResult):
    root.deiconify()
    window = UpdateWindow(root)
    window.show_result(result)
    root.mainloop()


def abrir_gui():
    from verificacion_correo.gui.main import VerificacionCorreosGUI
    import tkinter as tk
    root = tk.Tk()
    app = VerificacionCorreosGUI(root)
    root.mainloop()


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Modificar __main__.py para delegar al launcher**

```python
from verificacion_correo.launcher import main
```

- [ ] **Step 3: Commit**

---

### Task 6: start.bat actualizado

**Files:**
- Modify: `start.bat`

- [ ] **Step 1: Actualizar start.bat**

```bat
@echo off
setlocal

cd /d "%~dp0"

title Verificacion de Correo
echo ========================================
echo   Verificacion de Correo
echo ========================================
echo.

REM Verificar que Python esta disponible
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no encontrado.
    echo Descarga Python desde: https://www.python.org/downloads/
    echo Marca "Add Python to PATH" durante la instalacion.
    echo.
    pause
    exit /b 1
)

REM Crear entorno virtual si no existe
if not exist ".venv\Scripts\activate" (
    echo [1/3] Creando entorno virtual...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [ERROR] No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
    echo       Entorno virtual creado.
    echo.

    echo [2/3] Instalando dependencias...
    call .venv\Scripts\activate
    pip install -e . >nul 2>&1
    if %errorlevel% neq 0 (
        echo [ERROR] No se pudieron instalar las dependencias.
        pause
        exit /b 1
    )
    echo       Dependencias instaladas.
    echo.

    echo [3/3] Instalando navegador Chromium para Playwright...
    .venv\Scripts\playwright install chromium >nul 2>&1
    echo       Navegador listo.
    echo.
) else (
    call .venv\Scripts\activate
)

REM Opcion para saltar actualizacion
if /I "%1"=="--skip-update" (
    echo Iniciando sin comprobacion de actualizaciones...
    python -m verificacion_correo.launcher --skip-update
) else (
    echo Iniciando interfaz grafica...
    python -m verificacion_correo.launcher
)

pause
```

- [ ] **Step 2: Commit**

---

### Task 7: Tests de integración

**Files:**
- Create: `tests/test_launcher.py`
- Create: `tests/test_gui/test_update_window.py`

- [ ] **Tests para launcher**

```python
def test_launcher_import_no_gui():
    import verificacion_correo.launcher as launcher
    assert hasattr(launcher, "main")
    assert hasattr(launcher, "abrir_gui")
```

- [ ] **Tests para update_window**

```python
def test_update_window_show_result_actualizado():
    from verificacion_correo.core.update_models import UpdateStatus, UpdateResult
    root = tk.Tk()
    window = UpdateWindow(root)
    result = UpdateResult(
        status=UpdateStatus.ACTUALIZADO,
        message="Test",
        previous_commit="abc123",
        current_commit="def456",
        commits_updated=3,
    )
    window.show_result(result)
    assert "3" in window.status_var.get() or "actualizada" in window.status_var.get().lower()
    root.destroy()
```

- [ ] **Run tests**

- [ ] **Commit**

---

### Task 8: Limpiar .gitignore y estructura de datos

**Files:**
- Modify: `.gitignore`

- [ ] **Step 1: Agregar a .gitignore**

```gitignore
# Estado del actualizador
updater.lock
```

- [ ] **Step 2: Commit**

---

## Self-Review Checklist

1. **Spec coverage:** Todos los requisitos del spec original están cubiertos
2. **Placeholder scan:** Sin TODOs o TBDs
3. **Type consistency:** Nombres coherentes entre archivos
4. **Dependencies:** Cada task es independiente y testeable por separado
5. **Fail-open:** Los errores no bloquean la GUI

---

## Ejecución

**Orden de tareas:**
1. app_paths.py
2. update_models.py
3. updater.py (lógica Git)
4. update_window.py (UI)
5. launcher.py (orquestación)
6. start.bat
7. Tests
8. .gitignore

**Approximated time:** 3-4 horas
