@echo off
echo Start Menu Shortcut Creator - Windows Executable Builder
echo ======================================================
echo.

REM Check if Python is installed
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/
    pause
    exit /b
)

echo Checking Python version...
python --version

echo.
echo Building executable with PyInstaller...
pyinstaller --onefile --windowed --icon=assets\app_icon.ico --name "Start_Menu_Shortcut_Creator" main.py

echo.
if exist dist\Start_Menu_Shortcut_Creator.exe (
    echo Build successful! The executable is located at:
    echo dist\Start_Menu_Shortcut_Creator.exe
) else (
    echo Build failed. Please check the output for errors.
)

pause
