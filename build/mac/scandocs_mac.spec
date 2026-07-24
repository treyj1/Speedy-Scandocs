# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for Speedy Scandocs — macOS
# Built automatically by GitHub Actions (.github/workflows/build.yml)
# or manually on a Mac: pyinstaller build/mac/scandocs_mac.spec --clean
#
# Tesseract binary + tessdata are copied into the app bundle so no
# separate Tesseract install is required by the end user.

import os

from PyInstaller.utils.hooks import collect_data_files

ROOT = os.path.abspath(os.path.join(SPECPATH, '..', '..'))

# ttkbootstrap ships its own runtime assets (icon fonts, theme images) as
# package data, not code — hiddenimports alone won't bundle these, and
# newer ttkbootstrap versions (2.x+) load an icon .ttf at style-init time,
# so a missing file here crashes the app on first launch.
ttkbootstrap_datas = collect_data_files('ttkbootstrap')

# Tesseract paths on the build machine (GitHub Actions macos-13 runner
# with 'brew install tesseract' already run, or a local Mac with Homebrew)
TESS_BINARY = "/opt/homebrew/bin/tesseract"       # Apple Silicon
TESS_ROOT   = "/opt/homebrew"
if not os.path.isfile(TESS_BINARY):
    TESS_BINARY = "/usr/local/bin/tesseract"      # Intel Mac
    TESS_ROOT   = "/usr/local"

# Only bundle English language data — keeps the DMG small
TESS_ENG = os.path.join(TESS_ROOT, "share", "tessdata", "eng.traineddata")
TESS_OSD = os.path.join(TESS_ROOT, "share", "tessdata", "osd.traineddata")

# Generate .icns from the PNG logo if it doesn't exist yet
ICON_PNG  = os.path.join(ROOT, 'assets', 'GDJ Logo.png')
ICON_ICNS = os.path.join(ROOT, 'assets', 'GDJ Logo.icns')
if not os.path.isfile(ICON_ICNS) and os.path.isfile(ICON_PNG):
    import subprocess, tempfile, shutil
    iconset = tempfile.mkdtemp(suffix='.iconset')
    try:
        from PIL import Image
        src = Image.open(ICON_PNG).convert('RGBA')
        for sz in [16, 32, 64, 128, 256, 512]:
            src.resize((sz, sz), Image.LANCZOS).save(
                os.path.join(iconset, f'icon_{sz}x{sz}.png'))
            src.resize((sz*2, sz*2), Image.LANCZOS).save(
                os.path.join(iconset, f'icon_{sz}x{sz}@2x.png'))
        subprocess.run(['iconutil', '-c', 'icns', iconset, '-o', ICON_ICNS], check=True)
    except Exception as e:
        print(f"Warning: could not generate .icns: {e}")
    finally:
        shutil.rmtree(iconset, ignore_errors=True)

a = Analysis(
    [os.path.join(ROOT, 'scandocs_tool.py')],
    pathex=[ROOT],
    binaries=[
        (TESS_BINARY, '.'),     # tesseract binary — PyInstaller will collect its dylibs too
    ],
    datas=[
        (os.path.join(ROOT, 'assets'),          'assets'),
        # Client list is intentionally NOT bundled — it's real PII and must
        # never ship in the public installer. The user points Settings at
        # their own client_list.txt at install time.
        (TESS_ENG, 'tessdata'),                 # English language data
        (TESS_OSD, 'tessdata'),                 # Script detection data
    ] + ttkbootstrap_datas,
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
    icon=ICON_ICNS if os.path.isfile(ICON_ICNS) else None,
    bundle_identifier='com.gdj.speedyscandocs',
    version='1.0',
    info_plist={
        'CFBundleDisplayName': 'Speedy Scandocs',
        'CFBundleShortVersionString': '1.0',
        'CFBundleIconFile': 'GDJ Logo.icns',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '11.0',
    },
)
