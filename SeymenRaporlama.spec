# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec — hem Windows hem macOS için ortak temel
# Windows : pyinstaller SeymenRaporlama.spec
# macOS   : pyinstaller SeymenRaporlama.spec

import sys, os
block_cipher = None

a = Analysis(
    ['seymen_raporlama.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets', 'assets'),          # ikon, görseller
    ],
    hiddenimports=[
        'customtkinter',
        'PIL', 'PIL._tkinter_finder',
        'pandas', 'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timedeltas',
        'openpyxl',
        'python_pptx', 'pptx',
        'xlsxwriter',
        'tkinter', 'tkinter.ttk',
        'babel.numbers',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'PyQt5', 'wx'],
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
    name='SeymenRaporlama',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                     # pencere uygulaması, terminal yok
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/logo.ico',
)

# macOS .app paketi
app = BUNDLE(
    exe,
    name='SeymenRaporlama.app',
    icon='assets/logo.ico',
    bundle_identifier='com.seymen.raporlama',
    info_plist={
        'CFBundleShortVersionString': '1.0.0',
        'CFBundleVersion':            '1.0.0',
        'NSHighResolutionCapable':    True,
        'LSMinimumSystemVersion':     '10.13.0',
    },
)
