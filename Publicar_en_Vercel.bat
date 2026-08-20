@echo off
chcp 65001 >nul
title Publicar Catálogo en Vercel - Importadora Rivero
cls

echo ====================================================================
echo      🚀 PUBLICADOR AUTOMÁTICO DE CATÁLOGO A VERCEL / GITHUB
echo                  Importadora Rivero - Bolivia
echo ====================================================================
echo.

cd /d "%~dp0"

echo [1/4] Verificando archivos del catálogo...
if exist "catalogos.html" (
    echo       ✓ Copiando 'catalogos.html' a 'index.html' para la web principal...
    copy /y "catalogos.html" "index.html" >nul
) else (
    echo       ⚠️ No se encontró 'catalogos.html'. Se publicarán los archivos existentes.
)

echo.
echo [2/4] Preparando cambios para Git...
git add index.html catalogos.html catalogos_desktop.html catalogos_mobile.html vercel.json generar_catalogo.py web_generator.py Logo*.png Logo*.webp Logo*.jpg 2>nul
git add -u 2>nul

echo.
echo [3/4] Guardando versión en el historial...
for /f "tokens=1-4 delims=/.- " %%a in ("%date%") do (
    set FECHA=%%a-%%b-%%c
)
git commit -m "Actualización del catálogo online para clientes [%FECHA%]" >nul 2>nul
if %errorlevel% equ 0 (
    echo       ✓ Nuevos cambios registrados correctamente.
) else (
    echo       ℹ️ No había cambios nuevos o ya estaban guardados.
)

echo.
echo [4/4] Subiendo a GitHub y actualizando Vercel en la nube...
git push origin main
if %errorlevel% equ 0 (
    echo.
    echo ====================================================================
    echo   🎉 ¡ÉXITO! Tu catálogo ha sido subido a GitHub y Vercel.
    echo   En unos 10-20 segundos estará disponible en vivo en tu enlace web.
    echo ====================================================================
) else (
    echo.
    echo ====================================================================
    echo   ❌ Ocurrió un inconveniente al conectar con GitHub / Vercel.
    echo   Por favor verifica tu conexión a internet o credenciales de Git.
    echo ====================================================================
)

echo.
pause
