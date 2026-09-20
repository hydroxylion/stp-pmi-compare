@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PYTHONIOENCODING=utf-8"
set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"
echo.
"%PY%" doctor.py %*
echo.
pause
