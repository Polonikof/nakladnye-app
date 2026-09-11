@echo off
chcp 65001 >nul
cd /d "%~dp0"
python learn_novye.py --setup --folder "C:\Накладные\НовыеНакладные"
echo.
echo Папка: C:\Накладные\НовыеНакладные
echo Положите туда исходник и готовый файл, откройте Накладные.exe, нажмите «Обучить».
echo.
pause
