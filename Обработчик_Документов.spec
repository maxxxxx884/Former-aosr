# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['pretty_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('venv/Lib/site-packages/tkinterdnd2', 'tkinterdnd2')],
    hiddenimports=['PIL', 'fitz', 'tkinterdnd2', 'pythoncom', 'pywintypes'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [('O', None, 'OPTION'), ('O', None, 'OPTION')],
    name='Обработчик_Документов',
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
