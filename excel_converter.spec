# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for excel-converter — Windows standalone executable.
# Build with: python scripts/package.py  (or: pyinstaller excel_converter.spec --clean)

a = Analysis(
    ['src/excel_converter/__main__.py'],
    pathex=['src'],
    binaries=[],
    datas=[],
    hiddenimports=[
        # pywin32 COM modules — not auto-detected by PyInstaller's static analysis
        # because they are imported inside a try/except at runtime.
        'pythoncom',
        'pywintypes',
        'win32com',
        'win32com.client',
        'win32com.client.dynamic',
        'win32con',
        'win32api',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='excel-converter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
