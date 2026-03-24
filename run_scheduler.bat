@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo The virtual environment is not set up yet.
    echo Please double-click setup.bat first.
    echo.
    pause
    exit /b 1
)

call ".venv\Scripts\activate.bat"
python scheduler_gui.py

if errorlevel 1 (
    echo.
    echo The launcher closed with an error.
    pause
)
