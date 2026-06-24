# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
from shutil import copy2

a_fill = Analysis(
    ['src/autoexcel/main.py'],
    pathex=['src'],
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
a_diff = Analysis(
    ['src/autoexcel/diff_orders.py'],
    pathex=['src'],
    binaries=[],
    datas=[
        ('template/order_diff.html', 'template'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz_fill = PYZ(a_fill.pure)
pyz_diff = PYZ(a_diff.pure)

exe_fill = EXE(
    pyz_fill,
    a_fill.scripts,
    [],
    exclude_binaries=True,
    name='autoexcel-fill',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
exe_diff = EXE(
    pyz_diff,
    a_diff.scripts,
    [],
    exclude_binaries=True,
    name='diff-orders',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe_fill,
    exe_diff,
    a_fill.binaries,
    a_diff.binaries,
    a_fill.datas,
    a_diff.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='autoexcel',
)

Path(DISTPATH, 'autoexcel', 'workspace').mkdir(parents=True, exist_ok=True)
Path(DISTPATH, 'autoexcel', 'workspace', 'diffOrders').mkdir(parents=True, exist_ok=True)
Path(DISTPATH, 'autoexcel', 'result').mkdir(parents=True, exist_ok=True)
copy2('config.ini', Path(DISTPATH, 'autoexcel', 'config.ini'))
