# -*- mode: python ; coding: utf-8 -*-

from __future__ import annotations

from pathlib import Path


def _target_dir(path: Path) -> str:
    """Return the PyInstaller target directory for the given relative path."""
    target = str(path)
    return target if target and target != "." else "."


project_root = Path(__file__).resolve().parent

datas = []


def include_resource(relative_path: str) -> None:
    """Collect files that must accompany the executable."""

    source = project_root / relative_path
    if not source.exists():
        return

    if source.is_dir():
        for item in source.rglob("*"):
            if item.is_file():
                relative_target = Path(relative_path) / item.relative_to(source)
                datas.append((str(item), _target_dir(relative_target.parent)))
    else:
        datas.append((str(source), _target_dir(Path(relative_path).parent)))


for resource in [
    "templates",
    "logic/legal_entities.json",
    "rateapp.ico",
]:
    include_resource(resource)


a = Analysis(
    ["main.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
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
    name="main",
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
    icon=str(project_root / "rateapp.ico") if (project_root / "rateapp.ico").exists() else None,
)
