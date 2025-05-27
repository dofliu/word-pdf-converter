# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src\\integrated_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['win32com', 'win32com.client', 'pythoncom', 'docx2pdf', 'reportlab.lib.units'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt6'],
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
    name='Word_PDF_Converter',
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
