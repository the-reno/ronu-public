@echo off
cd /d "%~dp0"
call _setup.bat
if errorlevel 1 exit /b 1
"%VENV_PYTHON%" app.py
if errorlevel 1 pause
