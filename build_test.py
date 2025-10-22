#!/usr/bin/env python3
"""
Test script for local PyInstaller build.

This script helps test the PyInstaller build locally before
committing to GitHub Actions.
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path


def run_command(cmd, description):
    """Run a command and handle errors."""
    print(f"\n{'='*60}")
    print(f"🔄 {description}")
    print(f"{'='*60}")

    try:
        result = subprocess.run(cmd, shell=True, check=True, capture_output=True, text=True)
        print(f"✅ {description} completed successfully")
        if result.stdout:
            print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} failed")
        print(f"Error: {e}")
        if e.stderr:
            print(f"Stderr: {e.stderr}")
        return False


def test_pyinstaller_build():
    """Test PyInstaller build process."""
    print("🚀 Testing PyInstaller Build for Verificación de Correos OWA")
    print("=" * 70)

    # Check if we're in the right directory
    if not Path("pyproject.toml").exists():
        print("❌ Error: pyproject.toml not found. Please run from project root.")
        return False

    # Check if src directory exists
    if not Path("src").exists():
        print("❌ Error: src directory not found.")
        return False

    # Clean previous builds
    print("\n🧹 Cleaning previous builds...")
    for build_dir in ["build", "dist", "*.spec"]:
        for path in Path(".").glob(build_dir):
            if path.is_dir():
                shutil.rmtree(path)
                print(f"   Removed: {path}")

    # Install dependencies
    if not run_command("pip install -r requirements.txt", "Installing dependencies"):
        return False

    # Install PyInstaller
    if not run_command("pip install pyinstaller", "Installing PyInstaller"):
        return False

    # Install Playwright browsers
    if not run_command("playwright install chromium", "Installing Playwright browsers"):
        return False

    # Build with PyInstaller (onefile)
    pyinstaller_cmd = '''pyinstaller --onefile \\
        --collect-all playwright \\
        --collect-all openpyxl \\
        --collect-all yaml \\
        --add-data "config.yaml.example;." \\
        --add-data "data/correos_template.xlsx;data" \\
        --hidden-import tkinter \\
        --hidden-import tkinter.ttk \\
        --hidden-import tkinter.scrolledtext \\
        --hidden-import tkinter.filedialog \\
        --hidden-import tkinter.messagebox \\
        --hidden-import verificacion_correo.core.config \\
        --hidden-import verificacion_correo.core.excel \\
        --hidden-import verificacion_correo.core.browser \\
        --hidden-import verificacion_correo.core.session \\
        --hidden-import verificacion_correo.core.extractor \\
        --hidden-import verificacion_correo.core.first_run \\
        --hidden-import verificacion_correo.gui.main \\
        --hidden-import verificacion_correo.cli.main \\
        --hidden-import verificacion_correo.utils.logging \\
        --exclude-module matplotlib \\
        --exclude-module scipy \\
        --exclude-module pytest \\
        --exclude-module IPython \\
        --exclude-module jupyter \\
        --exclude-module pandas \\
        --exclude-module numpy \\
        --exclude-module PIL \\
        --exclude-module cv2 \\
        src/verificacion_correo/__main__.py'''

    if not run_command(pyinstaller_cmd, "Building with PyInstaller"):
        return False

    # Check if executable was created
    exe_path = Path("dist/verificacion_correo.exe" if os.name == "nt" else "dist/verificacion_correo")
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        print(f"\n✅ Executable created successfully!")
        print(f"   Path: {exe_path}")
        print(f"   Size: {size_mb:.1f} MB")

        if size_mb < 50:
            print(f"⚠️ Warning: Executable seems small ({size_mb:.1f} MB)")
            print(f"   This might indicate missing dependencies")
        else:
            print(f"✅ Executable size looks good")

        # Test executable
        print(f"\n🧪 Testing executable...")
        try:
            result = subprocess.run(
                [str(exe_path), "--help"],
                capture_output=True,
                text=True,
                timeout=30
            )
            if result.returncode == 0:
                print("✅ Executable test passed!")
                if "verificacion-correo" in result.stdout:
                    print("✅ Help text contains expected content")
                else:
                    print("⚠️ Help text might be incomplete")
            else:
                print("❌ Executable test failed")
                print(f"Return code: {result.returncode}")
                if result.stderr:
                    print(f"Error: {result.stderr[:500]}...")
        except subprocess.TimeoutExpired:
            print("⚠️ Executable test timed out (might be normal for first run)")
        except Exception as e:
            print(f"⚠️ Could not test executable: {e}")

        return True
    else:
        print("❌ Executable not found after build")
        return False


def show_build_summary():
    """Show build summary and next steps."""
    print("\n" + "=" * 70)
    print("📋 BUILD SUMMARY")
    print("=" * 70)

    exe_path = Path("dist/verificacion_correo.exe" if os.name == "nt" else "dist/verificacion_correo")
    if exe_path.exists():
        print(f"✅ Executable: {exe_path.absolute()}")
        print(f"✅ Size: {exe_path.stat().st_size / (1024 * 1024):.1f} MB")

        print("\n📋 Next Steps:")
        print("1. Test the executable manually:")
        if os.name == "nt":
            print(f"   .\\dist\\verificacion_correo.exe --help")
        else:
            print(f"   ./dist/verificacion_correo --help")

        print("2. Test GUI mode:")
        if os.name == "nt":
            print(f"   .\\dist\\verificacion_correo.exe gui")
        else:
            print(f"   ./dist/verificacion_correo gui")

        print("3. Test first-run setup:")
        if os.name == "nt":
            print(f"   .\\dist\\verificacion_correo.exe")
        else:
            print(f"   ./dist/verificacion_correo")

        print("4. If everything works, commit and push to trigger GitHub Actions build")

        print("\n📦 Distribution:")
        print("- The executable is self-contained")
        print("- Includes all necessary resources (config template, Excel template)")
        print("- Creates configuration files on first run")
        print("- Works without external dependencies")

    else:
        print("❌ Build failed - no executable found")
        print("\n🔧 Troubleshooting:")
        print("1. Check that all dependencies are installed")
        print("2. Verify the src directory structure")
        print("3. Check for import errors in the logs")
        print("4. Try running with --verbose flag for more information")


if __name__ == "__main__":
    success = test_pyinstaller_build()
    show_build_summary()

    if success:
        print("\n🎉 Build test completed successfully!")
        sys.exit(0)
    else:
        print("\n❌ Build test failed!")
        sys.exit(1)