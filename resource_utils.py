"""Utility helpers for accessing bundled resources."""

from __future__ import annotations

from pathlib import Path
import sys


def resource_path(relative: str) -> Path:
    """Return absolute path to resource for dev and PyInstaller builds.

    Parameters
    ----------
    relative:
        Path to the resource relative to the project root or the temporary
        directory used by PyInstaller.
    """
    base_path = getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)
    return Path(base_path) / relative
