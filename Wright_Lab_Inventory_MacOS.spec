# Wright_Lab_Inventory.spec
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],   # entry point script
    pathex=[],
    binaries=[],
    datas=[
        ('microscope.icns', '.'),    # bundle the icon
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter.test',
        'unittest',
        'doctest',
        'test',
        'distutils',
        'idlelib',
        'turtledemo',
    ],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Wright_Lab_Inventory',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
    upx=True,
    console=False,
    icon='microscope.icns'   # macOS icon
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Wright_Lab_Inventory'
)

app = BUNDLE(
    coll,
    name='Wright_Lab_Inventory.app',
    icon='microscope.icns',
    bundle_identifier=None
)