import logging

logger = logging.getLogger(__name__)

try:
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover - module may be missing on non-Windows
    win32com = None  # type: ignore


def get_excel_app():
    """Return a fresh Excel COM application instance.

    Earlier versions of the application kept a single global Excel instance
    alive for the whole session.  In practice this caused odd behaviour after
    several exports (e.g. the next file taking forever to save) because Excel
    retained state between runs.  Creating a new process for every export keeps
    the environment clean and avoids those hangs.
    """

    if win32com is None:
        raise RuntimeError("win32com.client is not available")

    # ``DispatchEx`` creates a completely new Excel process every time which
    # avoids leaking state between exports.  Callers are responsible for
    # tweaking any performance related flags (``ScreenUpdating`` etc.).
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    return app


def close_excel_app(app):  # pragma: no cover - best effort cleanup
    """Terminate the Excel COM application returned by :func:`get_excel_app`."""

    try:
        app.Quit()
    except Exception:
        pass
