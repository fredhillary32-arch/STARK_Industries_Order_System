@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
call "%~dp0tools\setup_environment.bat" || (pause & exit /b 1)
".venv\Scripts\python.exe" -m stark_support.menu
echo.
pause
