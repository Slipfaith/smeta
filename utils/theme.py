"""Application-wide palette configuration helpers."""

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPalette

# Unified style for the custom title bar used by the legacy rates UI.
TITLE_BAR_STYLE = "background-color: #ffffff; color: black;"


def apply_theme(app):
    """Apply the unified application palette to ``app``."""
    palette = QPalette()

    # -------- Window colors --------
    palette.setColor(QPalette.Window, Qt.white)
    palette.setColor(QPalette.Base, Qt.white)
    palette.setColor(QPalette.AlternateBase, QColor(240, 240, 240))
    palette.setColor(QPalette.WindowText, Qt.black)

    # -------- Button colors --------
    palette.setColor(QPalette.Button, QColor(240, 240, 240))
    palette.setColor(QPalette.ButtonText, Qt.black)

    # -------- Text colors --------
    palette.setColor(QPalette.Text, Qt.black)
    palette.setColor(QPalette.ToolTipBase, Qt.black)
    palette.setColor(QPalette.ToolTipText, Qt.black)
    palette.setColor(QPalette.BrightText, Qt.red)
    palette.setColor(QPalette.Link, QColor(42, 130, 218))
    palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
    palette.setColor(QPalette.HighlightedText, Qt.white)
    palette.setColor(QPalette.PlaceholderText, QColor("#A0AEC0"))

    app.setPalette(palette)
