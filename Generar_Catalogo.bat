@echo off
title Generador de Catálogos - Importadora Rivero
cls
echo =======================================================
echo   GENERANDO CATÁLOGO DE IMPORTADORA RIVERO...
echo =======================================================
echo.
cd /d "%~dp0"
python web_generator.py
echo.
echo =======================================================
echo   Proceso finalizado.
echo =======================================================
pause
