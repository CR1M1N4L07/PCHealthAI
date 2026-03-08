# -*- mode: python ; coding: utf-8 -*-
# PCHealthAI.spec  —  PyInstaller build spec
# Run:  pyinstaller PCHealthAI.spec

import os
block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('config.json',         '.'),      # app config
        ('system_info.py',      '.'),      # backend modules
        ('update_manager.py',   '.'),
        ('diagnosis_engine.py', '.'),
    ],
    hiddenimports=[
        'customtkinter',
        'PIL',
        'PIL._tkinter_finder',
        'psutil',
        'wmi',
        'pythoncom',
        'win32api',
        'win32con',
        'win32com',
        'win32com.client',
        'pywintypes',
        'packaging',
        'packaging.version',
        'packaging.specifiers',
        'packaging.requirements',
        'urllib',
        'urllib.request',
        'urllib.parse',
        'urllib.error',
        'email',
        'email.mime',
        'html',
        'http',
        'http.client',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy',
        'tkinter.test', 'unittest',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PCHealthAI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,                      # compress with UPX if available (smaller .exe)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                 # no black console window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',               # <-- put your icon.ico next to this spec file
    uac_admin=True,                # request admin on launch (UAC prompt)
    version='version_info.txt',    # optional — remove this line if you skip version_info.txt
)
