# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for excel-converter-gui — Windows GUI executable (no console window).
# Build with: python scripts/package.py  (or: pyinstaller excel_converter_gui.spec --clean)

a = Analysis(
    ['src/excel_converter/gui.py'],
    pathex=['src'],
    binaries=[],
    datas=[],
    hiddenimports=[
        # pywin32 COM modules — not auto-detected because they are imported
        # inside a try/except at runtime in com_mode.py.
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
    name='excel-converter-gui',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,   # no black terminal window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
