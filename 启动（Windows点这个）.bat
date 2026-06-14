@echo off
chcp 65001 >/dev/null 2>&1
REM ============================================================
REM  Law Firm Legal Assessment Report Generator - Windows Launcher
REM  Double-click to start. Close this window to stop the server.
REM ============================================================

cd /d "%~dp0"
set "PROJECT_DIR=%~dp0"

echo ============================================================
echo   Law Firm Legal Assessment Report Generator
echo ============================================================
echo.

REM --- Check Python ---
set "PYTHON="
where python >/dev/null 2>&1
if %ERRORLEVEL% equ 0 (
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo [OK] Found %%i
    set "PYTHON=python"
)

where python3 >/dev/null 2>&1
if %ERRORLEVEL% equ 0 if not defined PYTHON (
    for /f "tokens=*" %%i in ('python3 --version 2^>^&1') do echo [OK] Found %%i
    set "PYTHON=python3"
)

if not defined PYTHON (
    echo [ERROR] Python not found. Please install Python 3.9+
    echo          Download: https://www.python.org/downloads/
    echo          Make sure to check "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)

REM --- Create venv if missing ---
set "VENV_DIR=%PROJECT_DIR%venv"
if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo.
    echo [INFO] Creating Python virtual environment...
    %PYTHON% -m venv "%VENV_DIR%"
    if %ERRORLEVEL% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created.
)

REM --- Activate venv ---
call "%VENV_DIR%\Scripts\activate.bat"
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to activate virtual environment.
    pause
    exit /b 1
)
echo [OK] Virtual environment activated.

REM --- Install dependencies ---
if exist "%PROJECT_DIR%webapp\requirements.txt" (
    echo.
    echo [INFO] Checking dependencies...
    "%VENV_DIR%\Scripts\python.exe" -m pip install -r "%PROJECT_DIR%webapp\requirements.txt" -q
    if %ERRORLEVEL% neq 0 (
        echo [WARN] Some dependencies may have failed to install.
        echo        Trying to continue anyway...
    ) else (
        echo [OK] Dependencies ready.
    )
)

REM --- Start server ---
echo.
echo ============================================================
echo   Starting server...
echo   Browser will open at http://localhost:8000
echo   Close this window to stop the server.
echo ============================================================
echo.

REM --- Open browser ---
start "" "http://localhost:8000"

REM --- Run FastAPI ---
cd /d "%PROJECT_DIR%webapp"
"%VENV_DIR%\Scripts\python.exe" main.py

pause
