# Основной стиль приложения: фон окна и базовые параметры общих виджетов
# - QMainWindow: светло-серый фон (#f5f5f5)
# - QGroupBox: серые рамки и скругления, отдельное оформление для drag&drop
# - QPushButton: зелёная кнопка (#059669) с состояниями hover/pressed/disabled
# - Поля ввода и таблицы: светлые рамки, скругления и подсказки
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


# Специальные стили кнопок, используемые в виджете ставок
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


# Базовый стиль зоны перетаскивания: светло-серая пунктирная рамка и фон
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


# Стиль зоны перетаскивания в момент drag&drop: синяя пунктирная рамка и фон
DROP_AREA_DRAG_ONLY_STYLE = """
QScrollArea[dragOver="true"] {
    border: 2px dashed #2563eb;
    background-color: #eff6ff;
}
"""


# Подсказка внутри зоны перетаскивания: серый текст и пунктирная рамка
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


# Стиль подсказки с итогами: компактный шрифт 12px и серый цвет текста
SUMMARY_HINT_LABEL_STYLE = "font-size: 12px; padding: 4px; color: #555;"


# Стиль отображения итоговой суммы: полужирный текст 14px и тёмно-серый цвет
TOTAL_LABEL_STYLE = "font-weight: bold; font-size: 14px; padding: 6px; color: #333;"


# Стиль вспомогательных подписей к отчётам: маленький текст 11px серого цвета
REPORTS_LABEL_STYLE = "color: #555; font-size: 11px;"


# Стиль диалога импорта расценок: белый фон, синие акценты и таблицы
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


# Начальный размер окна расценок: ширина 1400 px, высота 720 px
RATES_WINDOW_INITIAL_SIZE = (1400, 720)

# Отступы главного layout окна расценок: 10 px со всех сторон
RATES_WINDOW_LAYOUT_MARGINS = (10, 10, 10, 10)

# Расстояние между элементами главного layout окна расценок
RATES_WINDOW_LAYOUT_SPACING = 0

# Коэффициенты растяжения сплиттера: левая панель 1, правая 2
RATES_WINDOW_SPLITTER_STRETCH_FACTORS = (1, 2)

# Базовые размеры панелей сплиттера в пикселях: 420 и 980
RATES_WINDOW_SPLITTER_SIZES = [420, 980]

# Отступы layout области сопоставления: 12 px
RATES_MAPPING_LAYOUT_MARGINS = (12, 12, 12, 12)

# Расстояние между элементами layout сопоставления: 10 px
RATES_MAPPING_LAYOUT_SPACING = 10

# Отступы между элементами панели управления сопоставления: 8 px
RATES_MAPPING_CONTROLS_SPACING = 8

# Ширина комбобокса «Применить ко всем»: 110 px
RATES_MAPPING_APPLY_COMBO_WIDTH = 110

# Ширины колонок таблицы сопоставления в пикселях
RATES_MAPPING_TABLE_COLUMN_WIDTHS = [180, 180, 100]
# Высота строк таблицы панели ставок (примерно 12 мм при 96 DPI)
RATES_MAPPING_TABLE_ROW_HEIGHT = 48

# Цвет подписи статуса по умолчанию: серый #666666
STATUS_LABEL_DEFAULT_STYLE = "color: #666666;"


# Цвет подписи статуса при успехе: зелёный #107c10
STATUS_LABEL_SUCCESS_STYLE = "color: #107c10;"


# Цвет подписи статуса при ошибке: красный #d13438
STATUS_LABEL_ERROR_STYLE = "color: #d13438;"
