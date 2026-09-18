# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_data_files

datas = [
    ('3.ico', '.'),
    ('assets/app_icon.png', 'assets'),
    ('Montserrat-Bold.ttf', '.'),
    ('presets/thesis_strict.json', 'presets'),
]
binaries = []
hiddenimports = ['olefile', 'lxml', 'PIL']

datas += collect_data_files('docx')
for package in ('customtkinter', 'tkinterdnd2'):
    package_data, package_binaries, package_hiddenimports = collect_all(package)
    datas += package_data
    binaries += package_binaries
    hiddenimports += package_hiddenimports

a = Analysis(
    ['gui_main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['flask', 'requests'],
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
    name='Yu_Works',
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
    icon=['3.ico'],
)
