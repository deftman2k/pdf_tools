# -*- mode: python ; coding: utf-8 -*-
import sys
from pathlib import Path

block_cipher = None

a = Analysis(
    ['pdf_tools.py'],
    pathex=['.'],  # 또는 os.getcwd(), 절대 경로 문자열 가능
    binaries=[
        ('poppler/Library/bin/*.dll', 'poppler/Library/bin'),
        ('poppler/Library/bin/*.exe', 'poppler/Library/bin'),
    ],
    datas=[],
    hiddenimports=[
        'tkinter',
        'PyPDF2',
        'pdf2image',
        'PIL._imagingtk',
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

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='pdf_tools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # True로 바꾸면 콘솔창도 표시됨
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='pdf_tools'
)
