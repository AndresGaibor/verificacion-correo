# activa el entorno y ejecuta el script de Python
@echo off
call .venv\Scripts\activate
python -m verificacion_correo.gui.main
pause