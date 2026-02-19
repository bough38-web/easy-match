# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('seller_assets/user_manual_v1.0.0.html', 'seller_assets'), ('assets', 'assets'), ('presets.json', '.'), ('replacements.json', '.')]
binaries = []
hiddenimports = ['pandas', 'xlwings', 'openpyxl', 'xlsxwriter', 'requests', 'PIL', 'PIL.Image', 'PIL.ImageTk', 'rapidfuzz', 'calamine', 'tkinterdnd2', 'excel_io', 'excel_io_additions']
tmp_ret = collect_all('Pillow')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5', 'PyQt6', 'qtpy', 'QtPy', 'jupyter', 'notebook', 'scipy', 'matplotlib', 'IPython', 'sympy', 'astropy'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ExcelMatcher_v1.0.0',
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
    name='ExcelMatcher_v1.0.0',
)
app = BUNDLE(
    coll,
    name='ExcelMatcher_v1.0.0.app',
    icon=None,
    bundle_identifier=None,
)
