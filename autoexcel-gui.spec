# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path
import re


APP_VERSION = re.search(
    r'^VERSION\s*=\s*"([^"]+)"',
    Path("src/autoexcel/version.py").read_text(encoding="utf-8"),
    flags=re.MULTILINE,
).group(1)


a = Analysis(
    ["src/autoexcel/gui.py"],
    pathex=["src"],
    binaries=[],
    datas=[
        ("template/order_diff.html", "template"),
        ("icon/cover-v4.png", "icon"),
        ("icon/sidebar", "icon/sidebar"),
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
    icon="icon/cover-v4.ico" if sys.platform == "win32" else None,
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
        icon="icon/cover-v4.icns",
        bundle_identifier="com.smartsheetdesk.desktop",
        info_plist={
            "CFBundleShortVersionString": APP_VERSION,
            "CFBundleVersion": APP_VERSION,
            "NSHighResolutionCapable": True,
        },
    )
