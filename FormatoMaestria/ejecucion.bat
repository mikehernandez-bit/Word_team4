@echo off
setlocal

cd /d "%~dp0"

set "DEFAULT_CONFIG=formats\unac_maestria_cuant.json"
set "CONFIG=%~1"
if "%CONFIG%"=="" set "CONFIG=%DEFAULT_CONFIG%"

where py >nul 2>nul
if not errorlevel 1 (
  py -3 -c "import docx" >nul 2>nul
  if errorlevel 1 (
    py -3 -m pip install --user python-docx
  )
  py -3 generate_from_json.py "%CONFIG%"
  goto end
)

where python >nul 2>nul
if not errorlevel 1 (
  python -c "import docx" >nul 2>nul
  if errorlevel 1 (
    python -m pip install --user python-docx
  )
  python generate_from_json.py "%CONFIG%"
  goto end
)

echo Python 3 not found on PATH.
echo Install Python and try again.

:end
pause
endlocal
