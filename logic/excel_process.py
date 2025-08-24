"""Utilities for managing Excel processes."""

from __future__ import annotations

import sys
from typing import Optional

try:
    import psutil  # type: ignore
except Exception:  # pragma: no cover - psutil might be missing
    psutil = None  # type: ignore


def close_excel_processes() -> None:
    """Attempt to terminate all running Excel processes on Windows.

    This helper is used as a safety net to ensure that stray Excel instances
    do not hang in memory if something goes wrong in the application.  It is
    a no-op on non-Windows systems or when :mod:`psutil` is not available.
    """
    if sys.platform != "win32" or psutil is None:
        return

    for proc in psutil.process_iter(["name"]):
        name: Optional[str] = proc.info.get("name")
        if name and name.lower() == "excel.exe":
            try:
                proc.terminate()
                proc.wait(timeout=5)
            except Exception:
                try:
                    proc.kill()
                except Exception:
                    pass
