"""Helpers for loading application configuration from ``.env`` files."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Iterable, Optional

from dotenv import load_dotenv

_ENV_FILENAME = ".env"
_loaded_path: Optional[Path] = None


def _candidate_directories() -> Iterable[Path]:
    """Yield directories that may contain the ``.env`` file."""

    cwd = Path.cwd()
    yield cwd

    # When running from source, the repository root is one level above this file.
    repo_root = Path(__file__).resolve().parent.parent
    yield repo_root

    # When running as a PyInstaller bundle ``sys._MEIPASS`` points to the
    # temporary extraction directory.
    bundle_dir = Path(getattr(sys, "_MEIPASS", repo_root))
    yield bundle_dir

    # When frozen, configuration is often next to the executable.
    if getattr(sys, "frozen", False):  # pragma: no cover - depends on runtime
        yield Path(sys.executable).resolve().parent

    custom_dir = os.getenv("SMETA_DOTENV_DIR")
    if custom_dir:
        yield Path(custom_dir)


def load_application_env() -> Optional[Path]:
    """Load the first available ``.env`` file from common locations."""

    global _loaded_path

    if _loaded_path is not None:
        return _loaded_path

    tried: set[Path] = set()
    for directory in _candidate_directories():
        path = Path(directory) / _ENV_FILENAME
        if path in tried:
            continue
        tried.add(path)
        if path.exists():
            load_dotenv(dotenv_path=path, override=False)
            _loaded_path = path
            return path

    # Fallback to default behaviour (load from current directory hierarchy).
    if load_dotenv(override=False):
        _loaded_path = Path(_ENV_FILENAME)
        return _loaded_path

    _loaded_path = None
    return None


__all__ = ["load_application_env"]

