@echo off
REM Arrastre el Excel de movimientos del BCP sobre este archivo.
chcp 65001 >nul
set PYTHONUTF8=1
if "%~1"=="" (
  echo Arrastre el Excel del BCP sobre este archivo.
  pause
  exit /b 1
)
python "%~dp0cargar_extracto.py" "%~1"
echo.
pause
