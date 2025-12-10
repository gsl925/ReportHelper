# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['report_helper_app_LocalUI_v15.3_advanced.py'],
    pathex=[],
    binaries=[],
    datas=[('D:\\_Personal\\_Coding\\_Python\\ReportHelper\\venv\\Lib\\site-packages\\tkinterdnd2', 'tkinterdnd2')],
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
    name='report_helper_app_LocalUI_v15.3_advanced',
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
    name='report_helper_app_LocalUI_v15.3_advanced',
)
