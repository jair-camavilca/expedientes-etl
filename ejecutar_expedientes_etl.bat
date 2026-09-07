@echo off
setlocal

set "PROJECT=%~dp0"
set "PYTHON_EXE="

where py >nul 2>&1
if not errorlevel 1 set "PYTHON_EXE=py -3"
if not defined PYTHON_EXE (
    where python >nul 2>&1
    if not errorlevel 1 set "PYTHON_EXE=python"
)

if not defined PYTHON_EXE (
    echo No se encontro Python. Instale Python 3 y vuelva a intentarlo.
    pause
    exit /b 1
)

cd /d "%PROJECT%"
echo Expedientes ETL
echo Entradas: %PROJECT%entrada
echo Salidas:  %PROJECT%salida
echo.
%PYTHON_EXE% expedientes_etl.py
set "EXIT_CODE=%ERRORLEVEL%"

if "%EXIT_CODE%"=="0" (
    echo.
    echo Proceso terminado correctamente.
    start "" explorer.exe "%PROJECT%salida"
) else (
    echo.
    echo El proceso termino con errores. Revise el mensaje anterior.
)

pause
exit /b %EXIT_CODE%
