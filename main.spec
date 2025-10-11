# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import os
import babel
import certifi
import langcodes
import language_data

block_cipher = None

# Безопасное определение корня проекта
project_root = Path(__file__).resolve().parent if '__file__' in globals() else Path.cwd()

# Дополнительные ресурсы
babel_locale_data_path = Path(babel.__file__).resolve().parent / "locale-data"
langcodes_data_path = Path(langcodes.__file__).resolve().parent / "data"

additional_datas = [
    (str(project_root / "templates" / "*"), "templates"),
    (str(project_root / "logic" / "legal_entities.json"), "logic"),
    (str(babel_locale_data_path), "babel/locale-data"),
    (str(langcodes_data_path), "langcodes/data"),
    (str(Path(language_data.__file__).resolve().parent / "data" / "*"), "language_data/data"),
    (certifi.where(), "certifi"),
    (str(project_root / "rateapp.ico"), "."),
]

a = Analysis(
    ['main.py'],
    pathex=[str(project_root)],
    binaries=[],
    datas=additional_datas,
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

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher,
)

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
    console=False,   # если нужно окно консоли -> True
    icon=str(project_root / 'rateapp.ico'),
)
