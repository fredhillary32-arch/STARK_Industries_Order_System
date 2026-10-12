@echo off
rem Creates the private Python environment in .venv the first time (needs Python 3.10 or newer),
rem and installs the packages in requirements.txt whenever one of them is missing.
set "VENV_PY=%~dp0..\.venv\Scripts\python.exe"
set "FRESH_VENV="
if exist "%VENV_PY%" goto packages
echo First run: setting up the Python environment. This takes a few minutes and only happens once.
where py >nul 2>nul
if not errorlevel 1 (
    py -3 -m venv "%~dp0..\.venv"
) else (
    python -m venv "%~dp0..\.venv"
)
if not exist "%VENV_PY%" (
    echo.
    echo Python was not found. Install Python 3.10 or newer from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during installation, then run this again.
    exit /b 1
)
set "FRESH_VENV=1"

:packages
"%VENV_PY%" -c "import openpyxl, streamlit" >nul 2>nul
if not errorlevel 1 exit /b 0
echo Installing the required packages. This takes a few minutes and only happens once.
"%VENV_PY%" -m pip install --disable-pip-version-check -q -r "%~dp0..\requirements.txt"
if not errorlevel 1 exit /b 0
echo.
echo Could not install the required packages. Check your internet connection and run this again.
if defined FRESH_VENV (
    rmdir /s /q "%~dp0..\.venv"
    exit /b 1
)
rem An existing environment is kept: the terminal menu only needs openpyxl.
"%VENV_PY%" -c "import openpyxl" >nul 2>nul
if errorlevel 1 exit /b 1
echo The terminal menu still works. The browser app needs the missing package.
exit /b 0
