"""Utilidades cross-platform para abrir archivos y carpetas.

Centraliza la lógica de apertura de archivos/carpetas en el explorador
del sistema, con manejo de errores y soporte para Windows, macOS y Linux.
"""

import os
import platform
import subprocess
from pathlib import Path
from subprocess import TimeoutExpired


def open_in_explorer(path: str | Path) -> bool:
    """Abrir archivo o carpeta en la aplicación predeterminada del sistema.

    Args:
        path: Ruta al archivo o directorio a abrir.

    Returns:
        True si se abrió correctamente, False si hubo error.
    """
    path = Path(path)

    # Si es directorio y no existe, crearlo primero
    if not path.exists():
        if path.suffix == '':
            # Parece un directorio (sin extensión)
            try:
                path.mkdir(parents=True, exist_ok=True)
            except OSError:
                return False
        else:
            # Es un archivo que no existe — no se puede abrir
            return False

    try:
        system = platform.system()
        if system == 'Windows':
            os.startfile(str(path))
            return True
        elif system == 'Darwin':
            result = subprocess.run(
                ['open', str(path)],
                capture_output=True,
                timeout=10,
            )
            return result.returncode == 0
        else:
            # Linux y otros
            result = subprocess.run(
                ['xdg-open', str(path)],
                capture_output=True,
                timeout=10,
            )
            return result.returncode == 0
    except FileNotFoundError:
        # Comando no encontrado (ej: xdg-open no instalado)
        return False
    except OSError:
        return False
    except TimeoutExpired:
        # Timeout — el proceso se lanzó pero no respondió
        return True  # Se lanzó correctamente, solo tardó


def open_file(path: str | Path) -> bool:
    """Abrir archivo con la aplicación predeterminada del sistema.

    Args:
        path: Ruta al archivo a abrir.

    Returns:
        True si se lanzó correctamente, False si hubo error.
    """
    path = Path(path)
    if not path.exists() or not path.is_file():
        return False
    return open_in_explorer(path)


def open_folder(folder: str | Path) -> bool:
    """Abrir carpeta en el explorador de archivos del sistema.

    Args:
        folder: Ruta al directorio a abrir.

    Returns:
        True si se lanzó correctamente, False si hubo error.
    """
    folder = Path(folder)
    if not folder.exists():
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except OSError:
            return False
    if not folder.is_dir():
        return False
    return open_in_explorer(folder)
