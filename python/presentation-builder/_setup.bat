@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" goto ready

where py >nul 2>nul
if not errorlevel 1 (
  py -3 -m venv .venv
) else (
  where python >nul 2>nul
  if errorlevel 1 goto nopython
  python -m venv .venv
)

if not exist ".venv\Scripts\python.exe" goto setupfailed
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto setupfailed

:ready
endlocal & set "VENV_PYTHON=%~dp0.venv\Scripts\python.exe"
exit /b 0

:nopython
echo Python 3 was not found.
echo Install it from https://www.python.org/downloads/windows/
pause
exit /b 1

:setupfailed
echo The Python environment could not be prepared.
pause
exit /b 1
