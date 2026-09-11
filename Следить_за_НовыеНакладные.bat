@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Папка: C:\Накладные\НовыеНакладные
echo Положите исходник и готовый — программа сама их разберёт.
echo Закройте окно, чтобы остановить.
echo.
python learn_novye.py --watch --folder "C:\Накладные\НовыеНакладные"
pause
