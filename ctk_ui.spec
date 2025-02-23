# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['ctk_ui.py','process_attendance_files.py','process_confirm_sheets.py','tools.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=["openpyxl.utils.dataframe","openpyxl.styles","openpyxl.utils","docx","docx.shared","docx.enum.text","docx.oxml.ns","docx2pdf"],
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
    name='上课啦数据处理',
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
    icon=['icon.ico'],
)
