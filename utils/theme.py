"""Application-wide palette and style definitions."""

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPalette

# Unified style for the custom title bar.
TITLE_BAR_STYLE = "background-color: #ffffff; color: black;"

# Style for "MLV_Rates_USD_EUR_RUR_CNY" button.
MLV_RATES_BUTTON_STYLE = """
QPushButton {
    background-color: #E2E8F0;
    border: 1px solid #CBD5E0;
    border-radius: 8px;
    padding: 6px 12px;
    font-size: 14px;
}
QPushButton:hover {
    background-color: #CBD5E0;
}
QPushButton:pressed {
    background-color: #A0AEC0;
}
"""

# Style for "TEP (Source RU)" button.
TEP_BUTTON_STYLE = """
QPushButton {
    background-color: #EDF2F7;
    border: 1px solid #D6BCFA;
    border-radius: 8px;
    padding: 6px 12px;
    font-size: 14px;
}
QPushButton:hover {
    background-color: #D6BCFA;
}
QPushButton:pressed {
    background-color: #B794F4;
}
"""


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
