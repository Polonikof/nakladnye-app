# -*- mode: python ; coding: utf-8 -*-
"""
Конфигурация PyInstaller: один исполняемый файл без установки.

Windows:  pyinstaller --noconfirm --clean nakladnye.spec   ->  dist/Накладные.exe
Linux:    pyinstaller --noconfirm --clean nakladnye.spec   ->  dist/Накладные

Справочник oboi_catalog.json вшивается внутрь. Если положить свою копию
рядом с исполняемым файлом, приложение возьмёт её (см. oboi_catalog_path).
"""

import os
import sys

BLOCK_CIPHER = None

# Рантайм Visual C++ лежит рядом с python.exe. Вшиваем его, чтобы .exe
# работал и на компьютерах без установленного «Microsoft VC++ Redistributable».
_extra_binaries = []
for _dll in ("vcruntime140.dll", "vcruntime140_1.dll"):
    _p = os.path.join(os.path.dirname(sys.executable), _dll)
    if os.path.exists(_p):
        _extra_binaries.append((_p, "."))

a = Analysis(
    ["nakladnye_app.py"],
    pathex=[],
    binaries=_extra_binaries,
    datas=[("oboi_catalog.json", "."), ("yusuf_profile.json", ".")],
    hiddenimports=["xlrd", "openpyxl", "learn_novye"],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        "numpy", "pandas", "matplotlib", "scipy", "PIL",
        "pytest", "setuptools", "pip", "sqlite3", "unittest",
    ],
    noarchive=False,
    cipher=BLOCK_CIPHER,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=BLOCK_CIPHER)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Накладные",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
