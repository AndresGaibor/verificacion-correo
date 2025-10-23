"""
Main entry point for verificacion-correo package.

This module allows the package to be executed with:
    python -m verificacion_correo
"""

import sys
from pathlib import Path

# Add src to path if running from development
if Path(__file__).parent.parent.name == "src":
    sys.path.insert(0, str(Path(__file__).parent.parent))

# Use absolute import instead of relative to avoid PyInstaller issues
from verificacion_correo.cli.main import main

if __name__ == "__main__":
    sys.exit(main())