"""Utilities for managing Excel processes."""

from __future__ import annotations

import sys
from contextlib import contextmanager
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
