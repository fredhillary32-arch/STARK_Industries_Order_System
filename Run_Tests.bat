@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
call "%~dp0tools\setup_environment.bat" || (pause & exit /b 1)
echo Running all tests. This takes about a minute.
".venv\Scripts\python.exe" tests\run_tests.py
echo.
pause
