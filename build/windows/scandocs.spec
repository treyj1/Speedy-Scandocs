# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for Speedy Scandocs — Windows
# Run from repo root:  pyinstaller build\windows\scandocs.spec --clean
#

import os

ROOT = os.path.abspath(os.path.join(SPECPATH, '..', '..'))

a = Analysis(
    [os.path.join(ROOT, 'scandocs_tool.py')],
    pathex=[ROOT],
    binaries=[],
    datas=[
        (os.path.join(ROOT, 'assets'),          'assets'),
        (os.path.join(ROOT, 'client_list.txt'), '.'),
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
