@echo off
chcp 65001 > nul
title Compilador - Convertidor XML a Excel IBC

echo.
echo ====================================================
echo   Compilador: Convertidor XML a Excel - IBC
echo ====================================================
echo.

:: Verificar Python
python --version > nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado o no esta en el PATH.
    echo Descargalo desde https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Instalar dependencias
echo Instalando dependencias...
python -m pip install --quiet --upgrade pyinstaller openpyxl
if errorlevel 1 (
    echo ERROR al instalar dependencias.
    pause
    exit /b 1
)

:: Llamar al script Python que hace la compilacion real
python "%~dp0_build_helper.py"
if errorlevel 1 (
    echo ERROR durante la compilacion.
    pause
    exit /b 1
)

pause
