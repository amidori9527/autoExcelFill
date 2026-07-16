# -*- mode: python ; coding: utf-8 -*-

import sys


a = Analysis(
    ["src/autoexcel/gui.py"],
    pathex=["src"],
    binaries=[],
    datas=[
        ("template/order_diff.html", "template"),
        ("config.ini", "."),
        ("loginConf.example.ini", "."),
        ("VERSION.txt", "."),
    ],
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
    name="SmartSheet Desk",
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
    name="SmartSheet Desk",
)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="SmartSheet Desk.app",
        icon="icon/cover.icns",
        bundle_identifier="com.smartsheetdesk.desktop",
        info_plist={"NSHighResolutionCapable": True},
    )
