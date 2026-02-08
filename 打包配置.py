# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['主程序.py'],
    pathex=[],
    binaries=[],
    datas=[('template.xlsx', '.')],
    hiddenimports=['pptx', 'pptx.enum.text', 'pptx.util', 'openpyxl', 'tkinter', 'tkinter.filedialog', 'tkinter.ttk', 'tkinter.messagebox', 'tkinter.scrolledtext'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'matplotlib', 'scipy', 'IPython', 'pytest', 'jinja2', 'markupsafe', 'pygments', 'tornado', 'sqlalchemy', 'psutil', 'lxml.isoschematron', 'lxml.objectify'],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='word_ppt_generator_light',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    onefile=True,
)
