@echo off
title Automatizacion de Facturas - FrasCL
echo ==========================================
echo    INICIANDO PROCESO DE FACTURACION
echo ==========================================
echo.
echo 1. Verificando Python...
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta instalado en este equipo.
    echo Por favor, contacta con sistemas o instala Python 3.12.
    pause
    exit /b
)

echo 2. Iniciando el script...
python generar_facturas.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Hubo un problema al generar los informes.
    echo Asegurate de que el Excel no este abierto y tenga el nombre correcto.
) else (
    echo.
    echo [EXITO] Informes generados correctamente en la carpeta 'PROYECTO FrasCL'.
)

echo.
echo Presiona cualquier tecla para cerrar esta ventana...
pause >nul
