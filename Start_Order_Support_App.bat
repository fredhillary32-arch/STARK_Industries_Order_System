@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
call "%~dp0tools\setup_environment.bat" || (pause & exit /b 1)
".venv\Scripts\python.exe" -c "import streamlit" >nul 2>nul
if errorlevel 1 (
    echo The browser app is not installed. Use Start_STARK_Order_System.bat for the terminal menu instead.
    pause
    exit /b 1
)
echo STARK INDUSTRIES Order Support is starting. Your browser opens at http://localhost:8501
echo Keep this window open while you work. To stop the app, close this window or press Ctrl+C.
echo Closing the browser tab does not stop it.
echo.
start "" /b ".venv\Scripts\python.exe" tools\open_browser_when_ready.py 8501
".venv\Scripts\python.exe" -m streamlit run app.py --server.address localhost --server.port 8501 --server.headless true --browser.gatherUsageStats false
echo.
pause
