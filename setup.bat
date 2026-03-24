@echo off
setlocal

cd /d "%~dp0"

echo ==========================================
echo Setting up Student Visit Scheduler...
echo ==========================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=python"
    ) else (
        echo Python was not found on this computer.
        echo Please install Python 3 first, then run this file again.
        echo.
        pause
        exit /b 1
    )
)

if not exist ".venv" (
    echo Creating virtual environment...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
) else (
    echo Virtual environment already exists.
)

echo.
echo Activating virtual environment...
call ".venv\Scripts\activate.bat"

echo.
echo Installing required packages...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo Package installation failed.
    pause
    exit /b 1
)

echo.
echo Setup complete.
echo You can now double-click run_scheduler.bat
echo.
pause
