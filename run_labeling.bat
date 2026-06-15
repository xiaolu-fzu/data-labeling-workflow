@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ================================================
echo   Data Labeling Workflow v3.0
echo ================================================
echo.
echo Press any key to start...
pause >nul
python data_labeling.py
echo.
echo Press any key to exit...
pause >nul
