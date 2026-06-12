# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['E:\\dev\\excel_tool\\Excel批量解密工具与密码管理.py'],
    pathex=[],
    binaries=[],
    datas=[('E:\\dev\\excel_tool\\requirements.txt', '.')],
    hiddenimports=['win32com', 'win32com.client', 'pythoncom', 'PyQt5', 'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'pandas', 'numpy', 'pytest'],
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
    name='Excel批量加密解密工具_vunknown',
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
    icon=['E:\\dev\\excel_tool\\图标.ico'],
)
