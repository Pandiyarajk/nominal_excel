@echo off
REM Quick start script for Excel Viewer & Editor Pro (Windows)

echo 🚀 Starting Excel Viewer & Editor Pro...
echo ================================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python is not installed or not in PATH
    echo Please install Python from https://python.org
    pause
    exit /b 1
)

REM Check if streamlit is installed
python -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo ⚠️  Streamlit is not installed.
    echo 📦 Installing required packages...
    echo.
    python -m pip install -q -r requirements.txt
    if errorlevel 1 (
        echo ❌ Failed to install requirements
        pause
        exit /b 1
    )
    echo ✅ Installation complete!
    echo.
)

echo 🌐 Starting application...
echo 📊 Open your browser at: http://localhost:8501
echo.
echo Press Ctrl+C to stop the server
echo ================================================
echo.

REM Run streamlit with custom config
python -m streamlit run excel_app.py --server.headless=true --server.address=0.0.0.0 --server.port=8501 --browser.gatherUsageStats=false

pause
