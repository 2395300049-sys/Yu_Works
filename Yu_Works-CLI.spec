# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['format_conversion.py'],
    pathex=[],
    binaries=[],
    datas=[('presets/thesis_strict.json', 'presets')],
    hiddenimports=['olefile', 'lxml'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['flask', 'requests', 'customtkinter'],
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
    name='Yu_Works-CLI',
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
    icon=['3.ico'],
)
