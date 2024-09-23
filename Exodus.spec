# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\Amitpate\\PycharmProjects\\Exodus\\Exodus.py'],
    pathex=[],
    binaries=[],
    datas=[('RobotSplash.png', '.'), ('Robot.ico', '.'), ('Manual.PDF', '.'), ('Changelog.txt', '.'), ('MAGNA.png', '.'), ('MAGNA.ico', '.'), ('MAGNAlogo.png', '.')],
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
    a.binaries,
    a.datas,
    [],
    name='Exodus',
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
    icon=['C:\\Users\\AmitPate\\PycharmProjects\\Exodus\\Robot.ico'],
)
