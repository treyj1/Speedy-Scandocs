# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for Speedy Scandocs — macOS
# Built automatically by GitHub Actions (.github/workflows/build.yml)
# or manually on a Mac: pyinstaller build/mac/scandocs_mac.spec --clean
#
# Tesseract binary + tessdata are copied into the app bundle so no
# separate Tesseract install is required by the end user.

import os

ROOT = os.path.abspath(os.path.join(SPECPATH, '..', '..'))

# Tesseract paths on the build machine (GitHub Actions macos-13 runner
# with 'brew install tesseract' already run, or a local Mac with Homebrew)
TESS_BINARY   = "/opt/homebrew/bin/tesseract"   # Apple Silicon
TESS_TESSDATA = "/opt/homebrew/share/tessdata"
if not os.path.isfile(TESS_BINARY):
    TESS_BINARY   = "/usr/local/bin/tesseract"   # Intel Mac
    TESS_TESSDATA = "/usr/local/share/tessdata"

a = Analysis(
    [os.path.join(ROOT, 'scandocs_tool.py')],
    pathex=[ROOT],
    binaries=[],
    datas=[
        (os.path.join(ROOT, 'assets'),          'assets'),
        (os.path.join(ROOT, 'client_list.txt'), '.'),
        (TESS_BINARY,   '.'),                   # tesseract binary -> bundle root
        (TESS_TESSDATA, 'tessdata'),            # language data
    ],
    hiddenimports=[
        'ttkbootstrap',
        'ttkbootstrap.themes',
        'ttkbootstrap.style',
        'PIL._tkinter_finder',
        'PIL.Image',
        'PIL.ImageTk',
        'fitz',
        'pytesseract',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'pkg_resources.py2_warn',
    ],
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
    [],
    exclude_binaries=True,
    name='SpeedyScandocs',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # UPX not recommended on macOS
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='SpeedyScandocs',
)

app = BUNDLE(
    coll,
    name='SpeedyScandocs.app',
    bundle_identifier='com.gdj.speedyscandocs',
    version='1.0',
    info_plist={
        'CFBundleDisplayName': 'Speedy Scandocs',
        'CFBundleShortVersionString': '1.0',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '11.0',
    },
)
