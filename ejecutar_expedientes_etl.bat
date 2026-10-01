@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "PROJECT=%~dp0"
set "PYTHON_EXE="

where py >nul 2>&1
if not errorlevel 1 (
    py -3 -c "import sys" >nul 2>&1
    if not errorlevel 1 set "PYTHON_EXE=py -3"
)
if not defined PYTHON_EXE (
    where python >nul 2>&1
    if not errorlevel 1 (
        python -c "import sys" >nul 2>&1
        if not errorlevel 1 set "PYTHON_EXE=python"
    )
)

if not defined PYTHON_EXE (
    echo No se encontro Python. Instale Python 3 y vuelva a intentarlo.
    pause
    exit /b 1
)

cd /d "%PROJECT%"
echo Expedientes ETL
echo Entrada:          %PROJECT%entrada
echo Salida:           %PROJECT%salida
echo Control calidad:  %PROJECT%salida\control_calidad
echo.

if not exist "%PROJECT%entrada\*.xlsx" if not exist "%PROJECT%entrada\*.xls" if not exist "%PROJECT%entrada\*.docx" if not exist "%PROJECT%entrada\*.doc" (
    echo Coloque los archivos Excel o Word en la carpeta entrada.
    start "" explorer.exe "%PROJECT%entrada"
    pause
    exit /b 0
)

%PYTHON_EXE% expedientes_etl.py
set "EXIT_CODE=%ERRORLEVEL%"

if "%EXIT_CODE%"=="0" (
    echo.
    echo Proceso terminado correctamente.
    echo Revise salida y salida\control_calidad.
    start "" explorer.exe "%PROJECT%salida"
    echo.
    set "SAVE_WORK="
    set /p "SAVE_WORK=Cuando termine de revisar, desea guardar cada archivo en trabajos? (S/N): "
    if /I "!SAVE_WORK!"=="S" (
        %PYTHON_EXE% expedientes_etl.py --guardar-trabajo
    ) else (
        echo El trabajo permanece en entrada y salida.
    )
) else (
    echo.
    echo El proceso termino con errores. Revise el mensaje anterior.
)

pause
exit /b %EXIT_CODE%
