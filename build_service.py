# build_service.py
import PyInstaller.__main__
import os
import sys

# Create the spec file content
spec_content = """# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['api_service.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'win32timezone',
        'flask',
        'waitress',
        'logging',
        'threading'
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
    name='APIService',
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
"""

# Write spec file
with open('api_service.spec', 'w') as f:
    f.write(spec_content)

# Run PyInstaller
if __name__ == "__main__":
    # Install required packages
    os.system('pip install pyinstaller pywin32 flask waitress')
    
    # Build the executable
    PyInstaller.__main__.run([
        'api_service.spec',
        '--clean',
        '--onefile'
    ])
