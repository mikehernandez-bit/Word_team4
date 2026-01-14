@echo off
setlocal
title Instalador y Ejecutor de Tesis UNAC
color 0A

:: Entrar a la carpeta donde esta el script
cd /d "%~dp0"

echo =====================================================
echo    SISTEMAS ALDAIR - INSTALADOR DE DEPENDENCIAS
echo =====================================================
echo.

:: Verificar Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta instalado o no se agrego al PATH.
    echo Por favor, instale Python desde python.org
    pause
    exit
)

:: Instalar python-docx si no existe
echo [*] Verificando libreria python-docx...
pip install python-docx --quiet

echo.
echo =====================================================
echo    GENERANDO INFORME DE TESIS... ESPERE...
echo =====================================================
echo.

:: Ejecutar el script
python Word.py

if %errorlevel% neq 0 (
    echo.
    echo [X] Error al generar el Word. 
    echo Verifique que el archivo no este abierto en este momento.
    pause
) else (
    echo.
    echo [OK] Documento generado y abierto con exito.
    timeout /t 5
)