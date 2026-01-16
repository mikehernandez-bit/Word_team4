@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

set "PYTHON_EXE=C:\Users\jhoan\AppData\Local\Python\pythoncore-3.14-64\python.exe"

if exist "%PYTHON_EXE%" (
  set "PYTHON_CMD=%PYTHON_EXE%"
) else (
  set "PYTHON_CMD=python"
)

echo Instalando dependencias desde requirements.txt...
"%PYTHON_CMD%" -m pip install -r "%SCRIPT_DIR%requirements.txt"

echo Iniciando servidor...
"%PYTHON_CMD%" "%SCRIPT_DIR%server.py"
