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
