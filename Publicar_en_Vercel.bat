@echo off
title Publicar Catalogo en Vercel - Importadora Rivero
cls

echo ====================================================================
echo      PUBLICADOR AUTOMATICO DE CATALOGO A VERCEL / GITHUB
echo                  Importadora Rivero - Bolivia
echo ====================================================================
echo.

cd /d "%~dp0"

echo [1/4] Verificando archivos del catalogo...
if exist "catalogos.html" (
    echo       Copiando catalogos.html a index.html para la web principal...
    copy /y "catalogos.html" "index.html" >nul
) else (
    echo       Aviso: No se encontro catalogos.html.
)

echo.
echo [2/4] Preparando cambios para Git...
git add index.html catalogos.html catalogos_desktop.html catalogos_mobile.html vercel.json generar_catalogo.py web_generator.py Publicar_en_Vercel.bat api/stock.js
git add -u

echo.
echo [3/4] Guardando version en el historial de Git...
git commit -m "Actualizacion del catalogo online para clientes"

echo.
echo [4/4] Subiendo a GitHub y actualizando Vercel en la nube...
git push origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ====================================================================
    echo   EXITO: Tu catalogo ha sido subido a GitHub y Vercel.
    echo   En unos 10-20 segundos estara disponible en vivo en tu enlace web.
    echo ====================================================================
) else (
    echo.
    echo ====================================================================
    echo   ERROR: Ocurrio un inconveniente al conectar con GitHub / Vercel.
    echo   Por favor verifica tu conexion a internet o credenciales de Git.
    echo ====================================================================
)

echo.
pause
