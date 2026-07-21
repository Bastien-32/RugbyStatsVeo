# -*- mode: python ; coding: utf-8 -*-

import sys


APPLICATION_NAME = "VeoVideoControl"
IS_MACOS = sys.platform == "darwin"
IS_WINDOWS = sys.platform == "win32"


a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[],
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
    name=APPLICATION_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=IS_WINDOWS,
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
    upx=IS_WINDOWS,
    upx_exclude=[],
    name=APPLICATION_NAME,
)


if IS_MACOS:
    app = BUNDLE(
        coll,
        name=f"{APPLICATION_NAME}.app",
        icon=None,
        bundle_identifier="fr.statsrugby.veovideocontrol",
        info_plist={
            "CFBundleName": APPLICATION_NAME,
            "CFBundleDisplayName": APPLICATION_NAME,
            "LSUIElement": True,
            "NSHighResolutionCapable": True,
        },
    )