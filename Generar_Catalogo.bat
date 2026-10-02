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

:: Verificar e instalar librerías automáticamente si faltan
python -c "import openpyxl, PIL" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [CONFIGURACIÓN] Verificando dependencias necesarias (Pillow, Openpyxl)...
    python -m pip install --quiet pillow openpyxl
    echo [CONFIGURACIÓN] Dependencias listas.
    echo.
)

python web_generator.py
echo.
echo =======================================================
echo   Proceso finalizado.
echo =======================================================
pause

