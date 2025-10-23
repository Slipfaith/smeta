# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import os

import babel
import certifi
import langcodes
import language_data
from PyInstaller.utils.hooks import collect_data_files


block_cipher = None

# Безопасное определение корня проекта
project_root = Path(__file__).resolve().parent if '__file__' in globals() else Path.cwd()


def collect_static_tree(path: Path, target: str) -> list[tuple[str, str]]:
    """Return a PyInstaller-style datas list for every file under *path*."""

    if not path.exists():
        return []

    entries: list[tuple[str, str]] = []
    target_path = Path(target) if target else Path(".")
    for file_path in path.rglob("*"):
        if file_path.is_file():
            relative = file_path.relative_to(path)
            destination_dir = target_path / relative.parent
            destination = str(destination_dir)
            if destination in {"", "."}:
                destination = "."
            entries.append((str(file_path), destination))
    return entries


# Дополнительные ресурсы
babel_locale_data_path = Path(babel.__file__).resolve().parent / "locale-data"
langcodes_data_path = Path(langcodes.__file__).resolve().parent / "data"
language_data_path = Path(language_data.__file__).resolve().parent / "data"

babel_locale_data = collect_data_files("babel", subdir="locale-data")
langcodes_data = collect_data_files("langcodes", subdir="data")
language_data_files = collect_data_files("language_data", subdir="data")
templates_data = collect_static_tree(project_root / "templates", "templates")

additional_datas = [
    (str(project_root / "logic" / "legal_entities.json"), "logic"),
    (str(babel_locale_data_path), "babel/locale-data"),
    (str(langcodes_data_path), "langcodes/data"),
    (str(language_data_path / "*"), "language_data/data"),
    (certifi.where(), "certifi"),
    (str(project_root / "rateapp.ico"), "."),
    *templates_data,
    *babel_locale_data,
    *langcodes_data,
    *language_data_files,
]


a = Analysis(
    ['main.py'],
    pathex=[str(project_root)],
    binaries=[],
    datas=additional_datas,
    hiddenimports=[
        'gui.styles',
        'utils.theme',
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
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,   # если нужно окно консоли -> True
    icon=str(project_root / 'rateapp.ico'),
)
