# -*- mode: python ; coding: utf-8 -*-

import sys
sys.path.append(SPECPATH)

from version import version_data

add_files = [("template\*", "template")]


a = Analysis(
    ["main.py"],  # 此列表存放项目设计的所有python脚本文件
    pathex=[],
    binaries=[],
    datas=add_files,  # 此列表存放所有资源文件，每个文件是一个2元组元素
    hiddenimports=["pkg_resources.extern"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="oa_tools-" + version_data["version"],
    version="build/win_version.py",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 启用cmd窗口，
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="assets/app_icon.ico",
)
