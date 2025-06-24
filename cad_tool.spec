# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets', 'assets'),
        ('config.py', '.'),
        ('utils', 'utils'),
        ('core', 'core'),
        ('ui', 'ui'),
    ],
    hiddenimports=[
        'pandas',
        'ezdxf',
        'openpyxl',
        'matplotlib',
        'PIL',
        'tkinter',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'tkinter.ttk',
        'utils.helpers',
        'utils.graphics',
        'core.rebar_processor',
        'core.excel_writer',
        'core.cad_reader',
        'core.dxf_parser',
        'ui.main_window',
        'ui.styles',
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
    exclude_binaries=True,  # 這行確保只產生資料夾模式
    name='CAD鋼筋計料轉換工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='assets/icon.ico'  # 暫時註解掉，因為檔案不存在
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CAD鋼筋計料轉換工具'
) 