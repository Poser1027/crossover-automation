@echo off
REM ─── Hengli Crossover launcher (Windows) ───
REM First-time setup will install dependencies.
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python not found. Please install Python 3.9+ from https://www.python.org/
    pause
    exit /b 1
)

REM Install deps quietly if missing
python -c "import fastapi, uvicorn, openpyxl" 2>nul
if errorlevel 1 (
    echo Installing dependencies...
    python -m pip install --quiet fastapi uvicorn openpyxl
)

echo.
echo ============================================================
echo  Hengli Crossover - running at http://127.0.0.1:8000
echo  Press Ctrl+C to stop.
echo ============================================================
echo.

REM Open browser automatically after 2 seconds
start "" /b cmd /c "timeout /t 2 /nobreak >nul && start http://127.0.0.1:8000"

python -m uvicorn api:app --host 127.0.0.1 --port 8000
