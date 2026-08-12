# -*- mode: python ; coding: utf-8 -*-
# Build on Windows:  build.bat   OR   pyinstaller PassPro2.spec
# Output: dist\PassPro2\PassPro2.exe  (run THAT file — not anything under build\)

from PyInstaller.building.api import COLLECT, EXE, PYZ
from PyInstaller.building.build_main import Analysis

datas = [
    ("assets/forest-dark.tcl", "assets"),
    ("assets/forest-light.tcl", "assets"),
    ("assets/passico.ico", "assets"),
    ("assets/passico.png", "assets"),
    ("assets/passico64.png", "assets"),
    ("assets/data2.template.xlsx", "assets"),
    ("assets/forest-dark", "assets/forest-dark"),
    ("assets/forest-light", "assets/forest-light"),
]

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name="PassPro2",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=["assets/passico.ico"],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="PassPro2",
)
