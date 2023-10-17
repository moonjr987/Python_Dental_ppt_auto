# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['ppt_auto_exe.py'],
    pathex=[],
    binaries=[],
    datas=[('C:/Users/ITL/anaconda3/envs/jy/Lib/site-packages/tkinterdnd2/tkdnd', '.'), ('C:/Users/ITL/anaconda3/envs/jy/Lib/site-packages/pptx/templates/default.pptx', 'pptx/templates'), ('123.png', '.'), ('aa.png', '.'), ('bb.png', '.'), ('cc.png', '.'), ('dd.png', '.'), ('ee.png', '.'), ('ee1.png', '.'), ('ee2.png', '.'), ('ee3.png', '.'), ('ee4.png', '.'), ('ee5.png', '.'), ('ee6.png', '.'), ('ee7.png', '.'), ('ee8.jpg', '.'), ('ff.png', '.'), ('gg.png', '.'), ('hh.png', '.'), ('ii.png', '.'), ('jj.png', '.')],
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
    name='ppt_auto_exe',
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
)
