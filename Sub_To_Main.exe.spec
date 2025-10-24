# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['sub_to_main.py'],
    pathex=[],
    binaries=[],
    datas=[('__init__.py', '.'), ('exiv2.dll', '.'), ('exiv2api.cpp', '.'), ('libexiv2.dylib', '.'), ('libexiv2.so', '.'), ('exiv2api.pyd', '.'), ('settings.json', '.')],
    hiddenimports=[],
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
    name='Sub_To_Main.exe',
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
    icon=['icon.ico'],
)
