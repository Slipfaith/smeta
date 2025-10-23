"""Application-wide palette configuration helpers."""

from PySide6.QtGui import QPalette

from gui.styles import (
    PALETTE_ALTERNATE_BASE_COLOR,
    PALETTE_BASE_COLOR,
    PALETTE_BRIGHT_TEXT_COLOR,
    PALETTE_BUTTON_COLOR,
    PALETTE_BUTTON_TEXT_COLOR,
    PALETTE_HIGHLIGHT_COLOR,
    PALETTE_HIGHLIGHTED_TEXT_COLOR,
    PALETTE_LINK_COLOR,
    PALETTE_PLACEHOLDER_TEXT_COLOR,
    PALETTE_TEXT_COLOR,
    PALETTE_TOOLTIP_BASE_COLOR,
    PALETTE_TOOLTIP_TEXT_COLOR,
    PALETTE_WINDOW_COLOR,
    PALETTE_WINDOW_TEXT_COLOR,
)


def apply_theme(app):
    """Apply the unified application palette to ``app``."""
    palette = QPalette()

    # -------- Window colors --------
    palette.setColor(QPalette.Window, PALETTE_WINDOW_COLOR)
    palette.setColor(QPalette.Base, PALETTE_BASE_COLOR)
    palette.setColor(QPalette.AlternateBase, PALETTE_ALTERNATE_BASE_COLOR)
    palette.setColor(QPalette.WindowText, PALETTE_WINDOW_TEXT_COLOR)

    # -------- Button colors --------
    palette.setColor(QPalette.Button, PALETTE_BUTTON_COLOR)
    palette.setColor(QPalette.ButtonText, PALETTE_BUTTON_TEXT_COLOR)

    # -------- Text colors --------
    palette.setColor(QPalette.Text, PALETTE_TEXT_COLOR)
    palette.setColor(QPalette.ToolTipBase, PALETTE_TOOLTIP_BASE_COLOR)
    palette.setColor(QPalette.ToolTipText, PALETTE_TOOLTIP_TEXT_COLOR)
    palette.setColor(QPalette.BrightText, PALETTE_BRIGHT_TEXT_COLOR)
    palette.setColor(QPalette.Link, PALETTE_LINK_COLOR)
    palette.setColor(QPalette.Highlight, PALETTE_HIGHLIGHT_COLOR)
    palette.setColor(QPalette.HighlightedText, PALETTE_HIGHLIGHTED_TEXT_COLOR)
    palette.setColor(QPalette.PlaceholderText, PALETTE_PLACEHOLDER_TEXT_COLOR)

    app.setPalette(palette)
