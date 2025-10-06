"""Utilities for managing Excel processes."""

from __future__ import annotations

import sys
from contextlib import contextmanager
from typing import Dict, Optional, Set

try:
    import psutil  # type: ignore
except Exception:  # pragma: no cover - psutil might be missing
    psutil = None  # type: ignore

# Track PIDs of Excel instances created by the application so we can
# gracefully shut them down when the program exits.
excel_instances: Set[int] = set()
_tracked_excel_objects: Dict[int, object] = {}


def _get_excel_pid(excel: object) -> Optional[int]:
    """Return the process ID of an Excel.Application COM object."""

    if sys.platform != "win32":
        return None

    try:
        hwnd = getattr(excel, "Hwnd")
    except Exception:
        return None

    try:
        import ctypes
        from ctypes import wintypes

        pid = wintypes.DWORD()
        result = ctypes.windll.user32.GetWindowThreadProcessId(  # type: ignore[attr-defined]
            hwnd, ctypes.byref(pid)
        )
        if result == 0:
            return None
        return int(pid.value)
    except Exception:  # pragma: no cover - Windows specific
        return None


def register_excel_instance(excel: object) -> None:
    """Register an Excel instance created by the application."""

    pid = _get_excel_pid(excel)
    if pid is None:
        return

    excel_instances.add(pid)
    _tracked_excel_objects[pid] = excel


def unregister_excel_instance(excel: object) -> None:
    """Remove an Excel instance from the registry if it is tracked."""

    pid = _get_excel_pid(excel)
    if pid is None:
        return

    excel_instances.discard(pid)
    _tracked_excel_objects.pop(pid, None)


def close_tracked_excel_instances() -> None:
    """Attempt to close only the Excel instances started by the app."""

    if sys.platform != "win32":
        excel_instances.clear()
        _tracked_excel_objects.clear()
        return

    for pid in list(excel_instances):
        excel = _tracked_excel_objects.get(pid)

        if psutil is not None:
            try:
                if not psutil.pid_exists(pid):
                    excel_instances.discard(pid)
                    _tracked_excel_objects.pop(pid, None)
                    continue
            except Exception:
                pass

        if excel is None:
            excel_instances.discard(pid)
            continue

        try:
            workbooks = getattr(excel, "Workbooks", None)
            if workbooks is not None:
                try:
                    for workbook in list(workbooks):  # type: ignore[arg-type]
                        try:
                            workbook.Close(False)
                        except Exception:
                            pass
                except TypeError:
                    try:
                        count = workbooks.Count  # type: ignore[attr-defined]
                    except Exception:
                        count = 0
                    for index in range(1, count + 1):
                        try:
                            workbooks(index).Close(False)  # type: ignore[call-arg]
                        except Exception:
                            pass
        except Exception:
            pass

        try:
            excel.Quit()  # type: ignore[attr-defined]
        except Exception:
            pass

        excel_instances.discard(pid)
        _tracked_excel_objects.pop(pid, None)


@contextmanager
def temporary_separators(excel, lang: str):
    """Temporarily override Excel decimal and thousands separators.

    Parameters
    ----------
    excel:
        ``win32com.client.Dispatch("Excel.Application")`` instance.
    lang: str
        Target language code (e.g. ``"en"`` or ``"ru"``).
    """
    orig_decimal = orig_thousands = orig_use_sys = None
    try:
        lang_lc = lang.lower()
        if lang_lc.startswith("en"):
            custom = (".", ",")
        else:
            custom = (",", " ")
        orig_decimal = excel.DecimalSeparator
        orig_thousands = excel.ThousandsSeparator
        orig_use_sys = excel.UseSystemSeparators
        excel.DecimalSeparator, excel.ThousandsSeparator = custom
        excel.UseSystemSeparators = False
        yield
    except Exception:
        # If anything goes wrong we still yield control so the caller can proceed
        yield
    finally:
        if orig_use_sys is not None:
            try:
                excel.DecimalSeparator = orig_decimal
                excel.ThousandsSeparator = orig_thousands
                excel.UseSystemSeparators = orig_use_sys
            except Exception:
                pass


def apply_separators(xlsx_path: str, lang: str) -> bool:
    """Open a workbook in Excel and save it with language-specific separators.

    This is primarily useful for ensuring that numbers appear with the desired
    decimal and thousands separators when the file is printed or exported to
    another format on a system with different locale settings.
    """
    excel = wb = None
    try:
        import win32com.client  # type: ignore

        excel = win32com.client.Dispatch("Excel.Application")
        register_excel_instance(excel)
        excel.Visible = False
        excel.DisplayAlerts = False

        with temporary_separators(excel, lang):
            wb = excel.Workbooks.Open(xlsx_path)
            wb.Save()
        return True
    except Exception:
        return False
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            unregister_excel_instance(excel)
