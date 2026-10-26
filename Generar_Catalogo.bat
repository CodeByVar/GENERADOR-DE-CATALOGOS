@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
title Generador de Catálogos - Importadora Rivero
cls
echo =======================================================
echo   GENERANDO CATÁLOGO DE IMPORTADORA RIVERO...
echo =======================================================
echo.
cd /d "%~dp0"

python web_generator.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Probando con comando py...
    py web_generator.py
)

echo.
echo =======================================================
echo   Proceso finalizado.
echo =======================================================
pause
