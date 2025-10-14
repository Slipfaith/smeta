APP_STYLE = """
QMainWindow {
    background-color: #f5f5f5;
}
QGroupBox {
    font-weight: bold;
    border: 1px solid #cccccc;
    border-radius: 5px;
    margin-top: 1ex;
    padding-top: 10px;
}
QGroupBox[dragOver="true"] {
    border: 2px dashed #2563eb;
    background-color: #eff6ff;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px 0 5px;
}
QPushButton {
    background-color: #059669;
    border: none;
    color: white;
    padding: 8px 16px;
    text-align: center;
    font-size: 14px;
    border-radius: 4px;
}
QPushButton:hover {
    background-color: #45a049;
}
QPushButton:pressed {
    background-color: #3d8b40;
}
QPushButton:disabled {
    background-color: #9ca3af;
    color: #f3f4f6;
}
QLineEdit, QTextEdit {
    border: 1px solid #ddd;
    border-radius: 4px;
    padding: 2px;
}
QTableWidget {
    border: 1px solid #ddd;
    border-radius: 4px;
}
QToolTip {
    color: #292524;
    background-color: #fef3c7;
    border: 1px solid #f59e0b;
    padding: 3px;
}
"""


DROP_AREA_BASE_STYLE = """
QScrollArea {
    border: 2px dashed #e5e7eb;
    border-radius: 8px;
    background-color: #fafafa;
}
QScrollArea[dragOver="true"] {
    border: 2px dashed #2563eb;
    background-color: #eff6ff;
}
"""


DROP_AREA_DRAG_ONLY_STYLE = """
QScrollArea[dragOver="true"] {
    border: 2px dashed #2563eb;
    background-color: #eff6ff;
}
"""


DROP_HINT_LABEL_STYLE = """
QLabel {
    color: #9ca3af;
    font-style: italic;
    padding: 24px;
    text-align: center;
    background-color: #f9fafb;
    border: 2px dashed #e5e7eb;
    border-radius: 8px;
    margin: 16px 0;
}
"""


SUMMARY_HINT_LABEL_STYLE = "font-size: 12px; padding: 4px; color: #555;"


TOTAL_LABEL_STYLE = "font-weight: bold; font-size: 14px; padding: 6px; color: #333;"


REPORTS_LABEL_STYLE = "color: #555; font-size: 11px;"


RATES_IMPORT_DIALOG_STYLE = """
QDialog {
    background-color: #ffffff;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 12px;
}

QLabel {
    color: #333333;
    font-weight: 500;
}

QLineEdit {
    padding: 4px 8px;
    border: 1px solid #cccccc;
    border-radius: 3px;
    background-color: white;
}

QLineEdit:focus {
    border-color: #0078d4;
}

QComboBox {
    padding: 4px 8px;
    border: 1px solid #cccccc;
    border-radius: 3px;
    background-color: white;
}

QComboBox:focus {
    border-color: #0078d4;
}

QComboBox::drop-down {
    border: none;
    width: 20px;
}

QComboBox::down-arrow {
    image: none;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 4px solid #666666;
    margin-right: 5px;
}

QPushButton {
    padding: 5px 12px;
    border: 1px solid #0078d4;
    border-radius: 3px;
    background-color: #0078d4;
    color: white;
    font-weight: 500;
}

QPushButton:hover {
    background-color: #106ebe;
    border-color: #106ebe;
}

QPushButton:pressed {
    background-color: #005a9e;
}

QTableWidget {
    gridline-color: #e1e1e1;
    background-color: white;
    alternate-background-color: #f8f8f8;
    border: 1px solid #cccccc;
    selection-background-color: #cce8ff;
}

QTableWidget::item {
    padding: 4px 6px;
    border: none;
}

QHeaderView::section {
    background-color: #f0f0f0;
    color: #333333;
    padding: 6px 8px;
    border: 1px solid #cccccc;
    border-left: none;
    font-weight: 600;
}

QHeaderView::section:first {
    border-left: 1px solid #cccccc;
}

QTableWidget QComboBox {
    border: none;
    padding: 2px 4px;
    margin: 1px;
}
"""


RATES_WINDOW_INITIAL_SIZE = (1400, 720)

RATES_WINDOW_LAYOUT_MARGINS = (10, 10, 10, 10)

RATES_WINDOW_LAYOUT_SPACING = 0

RATES_WINDOW_SPLITTER_STRETCH_FACTORS = (1, 2)

RATES_WINDOW_SPLITTER_SIZES = [420, 980]

RATES_MAPPING_LAYOUT_MARGINS = (12, 12, 12, 12)

RATES_MAPPING_LAYOUT_SPACING = 10

RATES_MAPPING_CONTROLS_SPACING = 8

RATES_MAPPING_APPLY_COMBO_WIDTH = 110

RATES_MAPPING_TABLE_COLUMN_WIDTHS = [140, 140, 160, 160, 80, 80, 80]

STATUS_LABEL_DEFAULT_STYLE = "color: #666666;"


STATUS_LABEL_SUCCESS_STYLE = "color: #107c10;"


STATUS_LABEL_ERROR_STYLE = "color: #d13438;"
