from PySide6.QtGui import QPalette, QColor
from PySide6.QtCore import Qt

# Style for the custom title bar when the light theme is active.
# "color" changes the text on the bar and "background-color" sets
# its background.
TITLE_BAR_STYLE = "background-color: #ffffff; color: black;"

# ------------------------------
# Styles for buttons used in RateTab when light theme is active.
# These constants mirror their dark theme counterparts but use
# a lighter color palette.
# ------------------------------

# Style for "MLV_Rates_USD_EUR_RUR_CNY" button
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

# Style for "TEP (Source RU)" button
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


def apply_light_theme(app):
    """Apply a light palette to the QApplication."""
    palette = QPalette()

    # -------- Window colors --------
    palette.setColor(QPalette.Window, Qt.white)  # Background of the main window
    palette.setColor(QPalette.Base, Qt.white)   # Background for text entry widgets
    palette.setColor(QPalette.AlternateBase, QColor(240, 240, 240))  # Alternate row colors
    palette.setColor(QPalette.WindowText, Qt.black)  # Text color for window elements and title bar

    # -------- Button colors --------
    palette.setColor(QPalette.Button, QColor(240, 240, 240))  # Button background
    palette.setColor(QPalette.ButtonText, Qt.black)           # Button text color

    # -------- Text colors --------
    palette.setColor(QPalette.Text, Qt.black)                 # Default text color
    palette.setColor(QPalette.ToolTipBase, Qt.black)          # Tooltip background
    palette.setColor(QPalette.ToolTipText, Qt.black)          # Tooltip text color
    palette.setColor(QPalette.BrightText, Qt.red)             # Bright text (warnings, etc.)
    palette.setColor(QPalette.Link, QColor(42, 130, 218))     # Hyperlink color
    palette.setColor(QPalette.Highlight, QColor(42, 130, 218))  # Selection background
    palette.setColor(QPalette.HighlightedText, Qt.white)      # Text color when selected
    palette.setColor(QPalette.PlaceholderText, QColor("#A0AEC0"))

    app.setPalette(palette)
