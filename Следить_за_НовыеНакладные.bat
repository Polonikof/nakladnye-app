@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Слежу за папкой НовыеНакладные.
echo Положите исходник и готовый файл — агент разберёт пару сам.
echo Закройте это окно, чтобы остановить наблюдение.
echo.
python learn_novye.py --watch
pause
