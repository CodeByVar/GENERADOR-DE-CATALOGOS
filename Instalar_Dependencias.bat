@echo off
chcp 65001 >nul
title Instalador de Dependencias
cls
echo =======================================================
echo   INSTALANDO DEPENDENCIAS DEL GENERADOR DE CATALOGOS
echo =======================================================
echo.
cd /d "%~dp0"

pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Probando con python -m pip...
    python -m pip install -r requirements.txt
)

echo.
echo =======================================================
echo   Proceso finalizado.
echo =======================================================
pause
