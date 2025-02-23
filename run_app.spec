# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import copy_metadata

datas = [
    ("E:/miniconda3/envs/pack/Lib/site-packages/streamlit/runtime", "./streamlit/runtime"),
    (os.path.join(os.getcwd(), "app.py"), "."),
    (os.path.join(os.getcwd(), "process_attendance_files.py"), "."),
    (os.path.join(os.getcwd(), "process_confirm_sheets.py"), "."),
    (os.path.join(os.getcwd(), "tools.py"), "."),
    (os.path.join(os.getcwd(), "icon.ico"), ".")

]
datas += collect_data_files("streamlit")
datas += copy_metadata("streamlit")

block_cipher = None


a = Analysis(
    ['run_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=["openpyxl.utils.dataframe","openpyxl.styles","openpyxl.utils","docx",
        "docx.shared","docx.enum.text","docx.oxml.ns","docx2pdf","tqdm"],
    hookspath=['./hooks'],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='上课啦表格制作',
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
    icon="icon.ico",
)
