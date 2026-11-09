@echo off
setlocal EnableDelayedExpansion
title Generador de Catalogos - Importadora Rivero
cls

echo =======================================================
echo      SISTEMA DE CATALOGOS - IMPORTADORA RIVERO
echo =======================================================
echo.

cd /d "%~dp0"

:: 1. Buscar el ejecutable de Python disponible
set "PY_CMD="

python --version >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=python"
) else (
    py --version >nul 2>&1
    if not errorlevel 1 (
        set "PY_CMD=py"
    ) else (
        if exist "%LOCALAPPDATA%\Python\bin\python.exe" (
            set "PY_CMD=%LOCALAPPDATA%\Python\bin\python.exe"
        ) else if exist "%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe" (
            set "PY_CMD=%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe"
        )
    )
)

if not defined PY_CMD (
    echo [ERROR] No se encontro Python instalado en el sistema.
    echo Asegurate de instalar Python y marcar la casilla 'Add Python to PATH'.
    echo.
    pause
    exit /b 1
)

:: 2. Verificar dependencias necesarias
%PY_CMD% -c "import openpyxl, PIL, svglib, reportlab" >nul 2>&1
if errorlevel 1 (
    echo [CONFIGURACION] Instalando dependencias necesarias: Pillow, openpyxl, svglib, reportlab...
    %PY_CMD% -m pip install --quiet pillow openpyxl svglib reportlab
    echo [CONFIGURACION] Dependencias listas.
    echo.
)

:: 3. Iniciar el servidor web interactivo
echo [INICIO] Abriendo el generador web en tu navegador...
echo [ENLACE] Local: http://localhost:5000
echo.
echo =======================================================
echo Presiona Ctrl+C en esta ventana para detener el servidor.
echo =======================================================
echo.

%PY_CMD% web_generator.py

echo.
echo =======================================================
echo   Servidor finalizado.
echo =======================================================
pause