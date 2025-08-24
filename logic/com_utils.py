import atexit
import logging

logger = logging.getLogger(__name__)

try:
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover - module may be missing on non-Windows
    win32com = None  # type: ignore

_excel_app = None


def get_excel_app():
    """Return a shared Excel COM application instance.

    Excel is started lazily and reused across multiple operations to avoid
    the heavy startup cost on each export.
    """
    global _excel_app
    if win32com is None:
        raise RuntimeError("win32com.client is not available")
    if _excel_app is None:
        _excel_app = win32com.client.Dispatch("Excel.Application")
        _excel_app.Visible = False
        _excel_app.DisplayAlerts = False
        atexit.register(_close_excel_app)
    return _excel_app


def _close_excel_app():  # pragma: no cover - best effort cleanup
    global _excel_app
    if _excel_app is not None:
        try:
            _excel_app.Quit()
        except Exception:
            pass
        _excel_app = None
