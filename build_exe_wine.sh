#!/usr/bin/env bash
# Кросс-сборка Windows-версии (Накладные.exe) на Linux через wine.
#
# Нужна, когда под рукой нет Windows: собирает тот же самый nakladnye.spec,
# но Windows-питоном под wine. Результат — dist-win/Накладные.exe,
# готовый к запуску на обычном ПК без установки Python.
#
# На Windows этот скрипт не нужен — там достаточно build_exe.bat.
set -euo pipefail

PYVER="${PYVER:-3.12.10}"
PYDIR="${PYDIR:-/tmp/winpy}"
export WINEPREFIX="${WINEPREFIX:-$HOME/.wine-py}"
export WINEDEBUG=-all
export DISPLAY="${DISPLAY:-:1}"

REPO="$(cd "$(dirname "$0")" && pwd)"
WINPY="$PYDIR/tools/python.exe"

say() { printf '\n=== %s ===\n' "$*"; }

say "Проверяю зависимости хоста"
for cmd in wine curl unzip msiextract; do
    command -v "$cmd" >/dev/null || {
        echo "Не найдено: $cmd"
        echo "Установите: sudo apt-get install -y wine64 curl unzip msitools"
        exit 1
    }
done

if [ ! -x "$WINPY" ]; then
    # Официальный установщик Python требует 32-битной прослойки wine, которой
    # обычно нет. Поэтому собираем Windows-Python из двух готовых архивов:
    #   • NuGet-пакет python — интерпретатор и стандартная библиотека;
    #   • tcltk.msi с python.org — tkinter, которого в NuGet-пакете нет.
    say "Собираю portable Windows-Python $PYVER"
    mkdir -p "$PYDIR"
    curl -fsSL --retry 3 -o "$PYDIR/python.nupkg" \
        "https://globalcdn.nuget.org/packages/python.$PYVER.nupkg"
    (cd "$PYDIR" && unzip -oq python.nupkg)

    tmp_tcltk="$(mktemp -d)"
    curl -fsSL --retry 3 -o "$tmp_tcltk/tcltk.msi" \
        "https://www.python.org/ftp/python/$PYVER/amd64/tcltk.msi"
    (cd "$tmp_tcltk" && msiextract tcltk.msi >/dev/null)
    cp -r "$tmp_tcltk/DLLs/." "$PYDIR/tools/DLLs/"
    cp -r "$tmp_tcltk/Lib/."  "$PYDIR/tools/Lib/"
    cp -r "$tmp_tcltk/tcl"    "$PYDIR/tools/"
    rm -rf "$tmp_tcltk"
fi

say "Проверяю Windows-Python"
wine "$WINPY" -c "import sys, tkinter; print(sys.version); print('tkinter', tkinter.TkVersion)"

say "Устанавливаю зависимости"
wine "$WINPY" -m pip install --no-warn-script-location -q \
    -r "$REPO/requirements.txt" -r "$REPO/requirements-build.txt"
wineserver -w

say "Собираю Накладные.exe"
cd "$REPO"
wine "$WINPY" -m PyInstaller --noconfirm --clean \
    --distpath dist-win --workpath build-win nakladnye.spec
wineserver -w

say "Готово"
ls -la "$REPO/dist-win/"
file "$REPO/dist-win/"*.exe
