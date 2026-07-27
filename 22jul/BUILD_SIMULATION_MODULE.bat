@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0assemble_simulation.ps1"
if errorlevel 1 (
    echo.
    echo Failed to create 02_Run_SOFR_Simulation.bas
    pause
    exit /b 1
)
echo.
echo 02_Run_SOFR_Simulation.bas is ready.
echo Import it together with 01_Build_SOFR_Template.bas into Excel VBA.
pause
