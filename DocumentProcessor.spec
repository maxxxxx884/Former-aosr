# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['pretty_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('path_manager.py', '.')],
    hiddenimports=['tkinterdnd2', 'docxtpl', 'docx2pdf', 'openpyxl', 'PyPDF2', 'fitz', 'PIL', 'pythoncom', 'appdirs', 'tqdm'],
    hookspath=['.'],
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
    name='DocumentProcessor',
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
