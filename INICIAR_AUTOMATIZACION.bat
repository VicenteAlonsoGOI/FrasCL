@echo off
title Automatizacion de Facturas - FrasCL
echo ==========================================
echo    INICIANDO PROCESO DE FACTURACION
echo ==========================================
echo.

:: Corregion para rutas de red (UNC)
:: pushd crea una unidad temporal si se ejecuta desde red
pushd "%~dp0"

echo 1. Verificando Python...
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta instalado en este equipo o no esta en el PATH.
    echo Por favor, contacta con sistemas o instala Python 3.12.
    popd
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
    echo [EXITO] Informes generados correctamente en la carpeta 'Facturas_Generadas'.
)

:: Quitar la unidad temporal
popd

echo.
echo Presiona cualquier tecla para cerrar esta ventana...
pause >nul
