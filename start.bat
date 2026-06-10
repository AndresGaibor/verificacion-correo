@echo off
REM Activa el entorno virtual y ejecuta la GUI clasica de verificacion de correos
call .venv\Scripts\activate
python iniciar_gui.py
pause
