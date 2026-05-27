@echo off
chcp 65001 >nul
title Auditoria de Riego por Equipo

:menu
cls
echo =============================================
echo   AUDITORIA DE RIEGO — TEMPORADA 2025-2026
echo =============================================
echo.
echo Equipos disponibles: 1 al 26
echo.
set /p equipo="Ingresa numero de equipo (Enter = salir): "

if "%equipo%"=="" goto fin

echo.
echo Procesando Equipo %equipo%...
echo.

python "%~dp0auditar_riegos.py" %equipo%

if %errorlevel% equ 0 (
    echo.
    echo Hecho. Abriendo Excel...
    start "" "%~dp0Auditoria_Riego_Equipo%equipo%.xlsx"
) else (
    echo.
    echo ERROR. Revisa que el archivo auditar_riegos.py exista.
)

echo.
set /p again="Queres procesar otro? (s/n): "
if /i "%again%"=="s" goto menu

:fin
echo.
echo Chau.
pause
