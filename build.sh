#!/usr/bin/env bash
# Сборка приложения в один файл (Linux / macOS).
# Результат: dist/Накладные
set -euo pipefail

cd "$(dirname "$0")"

python3 -m pip install -q -r requirements.txt
python3 -m pip install -q -r requirements-build.txt
python3 -m PyInstaller --noconfirm --clean nakladnye.spec

echo
echo "Готово: $(pwd)/dist/Накладные"
