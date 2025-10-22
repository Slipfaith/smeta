"""Helpers for ensuring Outlook COM cache availability."""

from __future__ import annotations

import csv
import logging
import shutil
import subprocess
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


def _get_outlook_process_ids() -> Optional[set[int]]:
    """Return the set of running Outlook process IDs, if discoverable."""

    try:
        completed = subprocess.run(
            [
                "tasklist",
                "/FI",
                "IMAGENAME eq OUTLOOK.EXE",
                "/FO",
                "CSV",
                "/NH",
            ],
            check=True,
            capture_output=True,
            text=True,
        )
    except Exception as exc:  # pragma: no cover - depends on local environment
        logger.debug("Failed to query Outlook processes: %s", exc)
        return None

    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    if not lines:
        return set()

    pids: set[int] = set()
    for line in lines:
        if line.startswith("INFO:"):
            return set()
        try:
            pid_str = next(csv.reader([line]))[1]
        except (IndexError, StopIteration):
            continue
        try:
            pids.add(int(pid_str))
        except ValueError:
            continue
    return pids


def _maybe_quit_outlook(app, started_new_instance: bool) -> None:
    """Close a temporary Outlook instance that was spawned for cache rebuild."""

    if not started_new_instance:
        return

    quit_method = getattr(app, "Quit", None)
    if quit_method is None:
        return
    try:
        quit_method()
    except Exception as exc:
        logger.debug("Failed to quit temporary Outlook instance: %s", exc)


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

    pythoncom_module = pythoncom
    coinitialized = False
    try:
        try:
            pythoncom_module.CoInitialize()
            coinitialized = True
        except Exception:
            # CoInitialize may fail if COM already initialised; continue anyway.
            coinitialized = False

        initial_pids = _get_outlook_process_ids()
        outlook_running_detected = bool(initial_pids)

        get_active = getattr(win32com.client, "GetActiveObject", None)
        if get_active is None:
            logger.debug(
                "win32com.client.GetActiveObject unavailable; Outlook running state unknown"
            )
        else:
            try:
                get_active("Outlook.Application")
            except Exception as exc:
                logger.debug("GetActiveObject for Outlook failed: %s", exc)
            else:
                outlook_running_detected = True

        def ensure_dispatch_and_cleanup() -> None:
            app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            latest_pids = _get_outlook_process_ids()

            started_new_instance = False
            if initial_pids is not None and latest_pids is not None:
                started_new_instance = bool(latest_pids - initial_pids)
            elif initial_pids is not None:
                started_new_instance = not initial_pids
            elif not outlook_running_detected:
                started_new_instance = False

            _maybe_quit_outlook(app, started_new_instance)

        try:
            ensure_dispatch_and_cleanup()
            return
        except Exception as exc:
            if outlook_running_detected:
                failure_note = "suspected cache error"
            elif initial_pids is not None:
                failure_note = "no Outlook process detected"
            else:
                failure_note = "unable to determine Outlook state"
            logger.warning(
                "Initial EnsureDispatch for Outlook failed (%s): %s; attempting to rebuild cache",
                failure_note,
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
            ensure_dispatch_and_cleanup()
        except Exception as exc:
            logger.error("Failed to rebuild win32com cache for Outlook: %s", exc)
        else:
            logger.info("win32com cache rebuilt successfully for Outlook")
    finally:
        if coinitialized:
            try:
                pythoncom_module.CoUninitialize()
            except Exception:
                pass
