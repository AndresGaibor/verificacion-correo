@echo off
REM Activa el entorno virtual y ejecuta la GUI de verificacion de correos
call .venv\Scripts\activate
python -m verificacion_correo.gui.main
pause
