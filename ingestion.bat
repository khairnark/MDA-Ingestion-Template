@echo off
cd /d "%~dp0"
cls

:repeat
echo ============================
echo   Running Ingestion Tool...
echo ============================

:: Create virtual environment if it does not exist
if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
)

:: Activate the virtual environment
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
) else (
    echo ERROR: activate.bat not found!
    pause
    exit /b
)

:: Install requirements only once
if not exist .requirements_installed (
    echo Installing required Python packages...
    pip install -r requirements.txt --quiet --disable-pip-version-check
    echo done > .requirements_installed
) else (
    echo Requirements already installed. Skipping installation.
)

:: Run the ingestion script
python ingestion.py

echo.
set /p choice=Do you want to run another ingestion? (Press ENTER to run again or type 'exit' to quit): 
if /I "%choice%"=="exit" goto end
goto repeat

:end
echo.
echo ============================
echo   Exiting Ingestion Tool.
echo ============================
pause
