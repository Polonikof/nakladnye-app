# -*- mode: python ; coding: utf-8 -*-
"""
Конфигурация PyInstaller: один исполняемый файл без установки.

Windows:  pyinstaller --noconfirm --clean nakladnye.spec   ->  dist/Накладные.exe
Linux:    pyinstaller --noconfirm --clean nakladnye.spec   ->  dist/Накладные

Справочник oboi_catalog.json вшивается внутрь. Если положить свою копию
рядом с исполняемым файлом, приложение возьмёт её (см. oboi_catalog_path).
"""

BLOCK_CIPHER = None

a = Analysis(
    ["nakladnye_app.py"],
    pathex=[],
    binaries=[],
    datas=[("oboi_catalog.json", ".")],
    hiddenimports=["xlrd", "openpyxl"],
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
