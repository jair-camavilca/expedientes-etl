@echo off
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" set "PYTHON_CMD=%~dp0.venv\Scripts\python.exe"

where py >nul 2>nul
if not defined PYTHON_CMD if not errorlevel 1 (
    py -3 -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=py"
        set "PYTHON_ARGS=-3"
    )
)

if not defined PYTHON_CMD (
    where python >nul 2>nul
    if not errorlevel 1 (
        python -c "import sys" >nul 2>nul
        if not errorlevel 1 (
            set "PYTHON_CMD=python"
            set "PYTHON_ARGS="
        )
    )
)

if not defined PYTHON_CMD (
    echo No se encontro Python funcional. Instale Python 3 y vuelva a intentarlo.
    pause
    exit /b 1
)

"%PYTHON_CMD%" %PYTHON_ARGS% -c "import streamlit" >nul 2>nul
if errorlevel 1 (
    echo Falta Streamlit. Ejecute: "%PYTHON_CMD%" %PYTHON_ARGS% -m pip install -r requirements.txt
    pause
    exit /b 1
)

"%PYTHON_CMD%" %PYTHON_ARGS% -m streamlit run app_web.py --server.headless true
pause
