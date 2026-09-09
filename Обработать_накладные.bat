@echo off
chcp 65001 >nul
cd /d "%~dp0"
python convert_nakladnaya.py
echo.
pause
