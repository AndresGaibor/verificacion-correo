"""
Configuracion de entorno para Windows y Mac/Linux.
Fuerza UTF-8 en stdout/stderr para evitar errores de encoding en cmd.exe
con caracteres como tildes, enies y simbolos.
"""
import sys
import os


def setup_console_encoding():
    """
    Configura la consola para usar UTF-8 en Windows.
    En Mac/Linux ya es UTF-8 por defecto, no hace nada.
    """
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass

        os.environ.setdefault("PYTHONIOENCODING", "utf-8")
        os.environ.setdefault("PYTHONUTF8", "1")
