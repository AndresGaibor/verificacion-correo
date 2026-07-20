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
