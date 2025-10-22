"""Helpers for ensuring Outlook COM cache availability."""

from __future__ import annotations

import logging
import shutil
import sys
from pathlib import Path
from typing import Optional


logger = logging.getLogger(__name__)


def _resolve_gen_py_path(win32com_module) -> Optional[Path]:
    """Return the ``gen_py`` directory for ``win32com`` if it exists."""

    try:
        gen_path = getattr(win32com_module, "__gen_path__")
    except AttributeError:
        try:
            module_path = Path(win32com_module.__file__).resolve()
        except Exception:
            return None
        gen_path = module_path.parent / "gen_py"
    try:
        path = Path(gen_path)
    except TypeError:
        return None
    return path


def rebuild_outlook_com_cache() -> None:
    """Ensure the Outlook COM cache is usable on Windows systems."""

    if sys.platform != "win32":
        return

    try:  # pragma: no cover - import itself is trivial but platform dependent
        import win32com  # type: ignore
        import win32com.client  # type: ignore
        import pythoncom  # type: ignore
    except Exception as exc:  # pragma: no cover - depends on local environment
        logger.debug("Skipping COM cache rebuild; pywin32 missing: %s", exc)
        return

    coinitialized = False
    try:
        try:
            pythoncom.CoInitialize()
            coinitialized = True
        except Exception:
            # CoInitialize may fail if COM already initialised; continue anyway.
            coinitialized = False

        try:
            win32com.client.gencache.EnsureDispatch("Outlook.Application")
            return
        except Exception as exc:
            logger.warning(
                "Initial EnsureDispatch for Outlook failed: %s; attempting to rebuild cache",
                exc,
            )

        gen_py_path = _resolve_gen_py_path(win32com)
        if gen_py_path and gen_py_path.exists():
            try:
                shutil.rmtree(gen_py_path)
                logger.info("Removed win32com gen_py cache at %s", gen_py_path)
            except Exception as cleanup_error:
                logger.error(
                    "Failed to clear win32com gen_py cache at %s: %s",
                    gen_py_path,
                    cleanup_error,
                )
                return
        else:
            logger.debug("win32com gen_py cache directory not found; skipping removal")

        try:
            win32com.client.gencache.EnsureDispatch("Outlook.Application")
        except Exception as exc:
            logger.error("Failed to rebuild win32com cache for Outlook: %s", exc)
        else:
            logger.info("win32com cache rebuilt successfully for Outlook")
    finally:
        if coinitialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
