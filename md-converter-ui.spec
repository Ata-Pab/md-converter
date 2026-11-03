# -*- mode: python ; coding: utf-8 -*-
# PyInstaller recipe for md-converter-ui.exe. Build it with build-exe.bat, or:
#     python -m PyInstaller --noconfirm md-converter-ui.spec
# Use "python -m PyInstaller", not the bare "pyinstaller" command: that one may
# belong to another Python installation that lacks the converter packages, and
# the .exe would then fail to convert those files.

import importlib.util
import sys

from PyInstaller.utils.hooks import collect_all, collect_data_files

REQUIRED = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "xlrd": "xlrd",
    "docx": "python-docx",
    "pptx": "python-pptx",
    "odf": "odfpy",
    "pdfplumber": "pdfplumber",
}
missing = [pkg for module, pkg in REQUIRED.items() if importlib.util.find_spec(module) is None]
if missing:
    raise SystemExit(
        "\nThis Python cannot build a working md-converter-ui.exe:\n"
        f"    {sys.executable}\n"
        f"is missing: {', '.join(missing)}\n"
        f"Install them:  \"{sys.executable}\" -m pip install {' '.join(missing)}\n"
        "or build with the Python that runs the app:  python -m PyInstaller md-converter-ui.spec\n"
    )

# Template / font-map files the libraries read at run time.
datas = []
for package in ("docx", "pptx", "pdfminer"):
    datas += collect_data_files(package)
binaries, hiddenimports = [], []
# Drag & drop is optional: bundled only when tkinterdnd2 is installed.
if importlib.util.find_spec("tkinterdnd2") is not None:
    dnd_datas, dnd_binaries, dnd_hidden = collect_all("tkinterdnd2")
    datas += dnd_datas
    binaries += dnd_binaries
    hiddenimports += dnd_hidden


a = Analysis(
    ['md-converter-ui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='md-converter-ui',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
