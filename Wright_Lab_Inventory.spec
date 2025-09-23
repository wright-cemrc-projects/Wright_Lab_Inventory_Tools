# Wright_Lab_Inventory.spec
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],   # entry point script
    pathex=[],
    binaries=[],
    datas=[
        ('microscope.ico', '.'),    # copy icon if app needs access at runtime
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
excludes=[
    'tkinter.test',   # GUI test suite
    'unittest',       # test framework
    'doctest',        # inline test runner
    'test',           # full Python test suite
    'distutils',      # rarely needed, unless you do builds at runtime
    'idlelib',        # IDLE editor
    'turtledemo',     # Turtle graphics demos
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
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
    strip=True,           # remove debug symbols for faster/smaller exe
    upx=True,             # compress binary (optional, test startup speed)
    console=False,        # no console window, GUI only
    icon='microscope.ico' # your .ico file
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