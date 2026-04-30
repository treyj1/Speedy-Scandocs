# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for Speedy Scandocs — Windows
# Run from repo root:  pyinstaller build\windows\scandocs.spec --clean
#

import os
import glob

import subprocess

ROOT = os.path.abspath(os.path.join(SPECPATH, '..', '..'))

# Find Tesseract — search common locations then fall back to PATH
_TESS_CANDIDATES = [
    r'C:\Program Files\Tesseract-OCR',
    r'C:\Program Files (x86)\Tesseract-OCR',
    r'C:\tools\tesseract',
    r'C:\ProgramData\chocolatey\lib\tesseract\tools',
]
TESS_ROOT = None
for _c in _TESS_CANDIDATES:
    if os.path.isfile(os.path.join(_c, 'tesseract.exe')):
        TESS_ROOT = _c
        break
if TESS_ROOT is None:
    try:
        _r = subprocess.run(['where', 'tesseract'], capture_output=True, text=True)
        if _r.returncode == 0:
            TESS_ROOT = os.path.dirname(_r.stdout.strip().splitlines()[0])
    except Exception:
        pass
if TESS_ROOT is None:
    raise RuntimeError("Tesseract not found. Install it before building.")
print(f"Bundling Tesseract from: {TESS_ROOT}")

# Bundle Tesseract: exe + all DLLs + English language data only
tess_datas = []
tess_datas += [(os.path.join(TESS_ROOT, 'tesseract.exe'), 'Tesseract-OCR')]
tess_datas += [(f, 'Tesseract-OCR') for f in glob.glob(os.path.join(TESS_ROOT, '*.dll'))]
tess_datas += [(os.path.join(TESS_ROOT, 'tessdata', 'eng.traineddata'), 'Tesseract-OCR/tessdata')]
tess_datas += [(os.path.join(TESS_ROOT, 'tessdata', 'osd.traineddata'), 'Tesseract-OCR/tessdata')]

a = Analysis(
    [os.path.join(ROOT, 'scandocs_tool.py')],
    pathex=[ROOT],
    binaries=[],
    datas=[
        (os.path.join(ROOT, 'assets'),          'assets'),
        # Client list is intentionally NOT bundled — it's real PII and must
        # never ship in the public installer. The user points Settings at
        # their own client_list.txt at install time.
    ] + tess_datas,
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
    upx=True,
    console=False,                          # no console window
    icon=os.path.join(ROOT, 'assets', 'GDJ Logo.ico'),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SpeedyScandocs',                  # output folder name in dist/
)
