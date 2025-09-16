# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['E:/PythonProjects/smeta'],
    binaries=[],
    datas=[
        # твои ресурсы
        ('E:/PythonProjects/smeta/languages/*', 'languages'),
        ('E:/PythonProjects/smeta/templates/*', 'templates'),
        ('E:/PythonProjects/smeta/logic/legal_entities.json', 'logic'),

        # Babel locale-data (чтобы не было ошибки "babel data files are not available")
        ('E:/PythonProjects/smeta/.venv/Lib/site-packages/babel/locale-data/*', 'babel/locale-data'),
    ],
    hiddenimports=[],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='RateApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # console=False == твой флаг --noconsole
    icon='E:/PythonProjects/smeta/rateapp.ico',
)