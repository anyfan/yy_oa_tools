# -*- mode: python ; coding: utf-8 -*-

add_files = [
    ('mb.docx', '.'),
    ('new.docx', '.')
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=add_files,
    hiddenimports=['pkg_resources.extern'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='main',
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
    icon='assets/app_icon.ico'
)
