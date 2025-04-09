@echo off
setlocal enabledelayedexpansion

:: Set project metadata
set APP_NAME=ReportConverter
set SRC_FILE=src\main.py
set TEMPLATE_DIR=report_templates

:: Generate timestamp for the output file (e.g., 2024-04-09_1532)
for /f %%a in ('powershell -Command "Get-Date -Format yyyy-MM-dd_HHmm"') do set TIMESTAMP=%%a
set OUTPUT_NAME=%APP_NAME%_%TIMESTAMP%.exe

echo.
echo === Building %APP_NAME% (%TIMESTAMP%) ===
echo.

:: Step 0: Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo ❌ Failed to activate virtual environment.
    pause
    exit /b 1
)

:: Step 1: Clean previous build artifacts
echo Cleaning previous build artifacts...
rmdir /s /q build >nul 2>&1
rmdir /s /q dist >nul 2>&1
del /q main.spec >nul 2>&1

:: Step 2: Run PyInstaller with timestamped name
echo.
echo Running PyInstaller...
pyinstaller --onefile --noconsole --name "%OUTPUT_NAME%" --add-data "%TEMPLATE_DIR%;%TEMPLATE_DIR%" %SRC_FILE%
if errorlevel 1 (
    echo.
    echo ❌ Build failed. Check for errors above.
    pause
    exit /b 1
)

:: Step 3: Notify and open dist folder
echo.
echo ✅ Build complete!
echo Executable created: dist\%OUTPUT_NAME%

start "" dist
echo.
pause
