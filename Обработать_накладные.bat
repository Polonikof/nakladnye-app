@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Обработка накладных (Витебск .xls и Обои УПД .xlsx)...
echo.
python convert_nakladnaya.py
echo.
pause
