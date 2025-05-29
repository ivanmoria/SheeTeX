# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['sheet.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PySide6', 'PySide6.QtCore', 'PySide6.QtWidgets', 'PySide6.QtGui'],
    noarchive=False,
    optimize=0,

    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,

)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='sheet',
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
    icon=['/imagens/we.icns']

)
app = BUNDLE(
    exe,
    name='sheet.app',
    icon='we.icns',
    bundle_identifier=None,
)
