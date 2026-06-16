@echo off
chcp 65001 >nul
echo ============================================
echo   Formula Converter - Setup
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found, please install Python 3.10+
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Create venv
if exist ".venv" (
    echo [SKIP] Virtual environment already exists
) else (
    echo [1/3] Creating virtual environment...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment
        pause
        exit /b 1
    )
    echo [DONE] Virtual environment created
)

:: Install dependencies
echo [2/3] Installing dependencies...
call .venv\Scripts\activate.bat
pip install -r requirements.txt -q
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)
echo [DONE] Dependencies installed

:: Verify
echo [3/3] Verifying installation...
python -c "import docx, lxml, latex2mathml" 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Dependency verification failed
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Setup complete!
echo   Run:     .venv\Scripts\python.exe main.py
echo   Build:   .venv\Scripts\python.exe build.py
echo ============================================
pause
