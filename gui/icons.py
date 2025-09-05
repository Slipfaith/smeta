"""Centralised icon definitions for the application.

All icons used across the GUI should be retrieved via :func:`get_icon` to
allow easy customisation. To replace an icon, specify a file path in the
``ICON_FILES`` mapping for the corresponding icon name.

If no custom path is provided or the file does not exist, a reasonable
``QStyle`` standard icon is used as a fallback.
"""

from __future__ import annotations

import os
from typing import Dict

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QStyle


# Map icon names to optional file paths. Users can edit this dictionary to
# point to their own icon files.
ICON_FILES: Dict[str, str] = {
    # "remove": "/path/to/remove_icon.png",
}

# Default fallback icons provided by QStyle.
DEFAULT_ICONS: Dict[str, QStyle.StandardPixmap] = {
    "remove": QStyle.SP_TrashIcon,
}


def get_icon(name: str) -> QIcon:
    """Return an icon by name.

    Parameters
    ----------
    name: str
        Identifier of the icon. Examples: ``"remove"``.

    Returns
    -------
    QIcon
        The icon loaded from ``ICON_FILES`` if present, otherwise the
        corresponding ``QStyle`` standard icon.
    """

    path = ICON_FILES.get(name)
    if path and os.path.exists(path):
        return QIcon(path)

    app = QApplication.instance()
    style = app.style() if app else None
    if style and name in DEFAULT_ICONS:
        return style.standardIcon(DEFAULT_ICONS[name])
    return QIcon()


__all__ = ["get_icon", "ICON_FILES", "DEFAULT_ICONS"]

