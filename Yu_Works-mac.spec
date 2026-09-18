# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

import docx
from PyInstaller.utils.hooks import collect_all, collect_data_files

datas = [
    ('Montserrat-Bold.ttf', '.'),
    ('3.ico', '.'),
    ('assets/app_icon.png', 'assets'),
    ('presets/thesis_strict.json', 'presets'),
]
binaries = []
datas += collect_data_files('docx')
# python-docx loads its default header/footer templates relative to
# ``docx/parts/hdrftr.py``.  In a macOS PyInstaller bundle the Python package
# itself is represented by a Frameworks symlink while package data lives in
# Resources.  Keep a real ``docx/parts`` directory in Resources so paths such
# as ``Frameworks/docx/parts/../templates/default-footer.xml`` can be resolved
# by the OS instead of failing before the ``..`` component is normalized.
_docx_package = Path(docx.__file__).resolve().parent
datas.append((str(_docx_package / 'parts' / '__init__.py'), 'docx/parts'))
hiddenimports = [
    'olefile',
    'lxml',
]
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
    [],
    exclude_binaries=True,
    name='Yu_Works',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Yu_Works',
)
app = BUNDLE(
    coll,
    name='Yu_Works.app',
    icon='assets/app_icon.icns',
    bundle_identifier='com.yuworks.desktop',
)
