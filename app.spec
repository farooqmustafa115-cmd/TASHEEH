# -*- mode: python ; coding: utf-8 -*-

import os

project_dir = os.path.abspath(SPECPATH)

a = Analysis(
    ['app.py'],
    pathex=[project_dir],
    binaries=[],
    datas=[
        (os.path.join(project_dir, 'شعار_الكلية_العسكرية_الأولى.png'), '.'),
        (os.path.join(project_dir, 'logo_watermark_transparent.gif'), '.'),
    ],
    hiddenimports=[
        'cv2',
        'openpyxl',
        'pymupdf',
        'pytesseract',
        'pyzbar',
        'qrcode.image.pil',
        'xlrd',
    ],
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
    name='app',
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
