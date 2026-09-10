@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo  Сборка приложения в один файл: dist\Накладные.exe
echo ============================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo Python не найден. Установите Python 3.9+ с python.org
    echo и при установке отметьте "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)

echo [1/2] Устанавливаю зависимости сборки...
python -m pip install --disable-pip-version-check -q -r requirements.txt
python -m pip install --disable-pip-version-check -q -r requirements-build.txt
if errorlevel 1 (
    echo Не удалось установить зависимости.
    pause
    exit /b 1
)

echo [2/2] Собираю исполняемый файл...
python -m PyInstaller --noconfirm --clean nakladnye.spec
if errorlevel 1 (
    echo Сборка не удалась.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Готово: "%~dp0dist\Накладные.exe"
echo.
echo  Скопируйте этот один файл в рабочую папку с накладными
echo  (например C:\Накладные) и запускайте двойным щелчком.
echo  Python на рабочем компьютере при этом не нужен.
echo ============================================================
echo.
pause
