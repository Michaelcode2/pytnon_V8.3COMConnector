# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Add the queries directory to the datas
added_files = [
    ('settings', 'settings')
]

a = Analysis(
    ['service.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        'win32timezone',
        'win32serviceutil',
        'win32service',
        'win32event',
        'servicemanager',
        'flask',
        'waitress',
        'logging',
        'threading',
        'win32com.client.dynamic',
        'pywintypes',
        'win32api'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PriceCheckerService',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
