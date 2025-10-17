# flake8: noqa

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor

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
    background-color: #374151;         /* slate-700 */
    color: #F9FAFB;                    /* почти белый */
    border: none;
    border-radius: 6px;
    padding: 6px 12px;
    font-size: 14px;

    /* Тень */
    box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.2);
}

QPushButton:hover {
    background-color: #4B5563;         /* slate-600 */
    box-shadow: 0px 4px 8px rgba(0, 0, 0, 0.3);  /* посильнее тень при наведении */
}

QPushButton:pressed {
    background-color: #1F2937;         /* slate-800 */
    padding-top: 7px;                  /* визуальный эффект "вжалось" */
    padding-bottom: 5px;
    box-shadow: inset 0px 2px 4px rgba(0, 0, 0, 0.4);  /* внутренняя тень */
}
"""

RATE_SELECTION_ACTION_BUTTON_STYLE = """
QPushButton {
    background-color: #e5e7eb;         /* slate-200 */
    color: #1f2937;                    /* slate-800 */
    border: 1px solid #d1d5db;         /* slate-300 */
    border-radius: 6px;
    padding: 6px 12px;
    font-size: 13px;
}

QPushButton:hover {
    background-color: #d1d5db;
}

QPushButton:pressed {
    background-color: #cbd5f5;
}

QPushButton:disabled {
    background-color: #f3f4f6;
    color: #9ca3af;
    border-color: #e5e7eb;
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


# Унифицированные размеры и отступы, используемые в нескольких модулях
ZERO_MARGINS = (0, 0, 0, 0)
GROUP_SECTION_MARGINS = (8, 4, 8, 4)
GROUP_SECTION_SPACING = 8


# Настройки кнопки удаления языковой пары
LANGUAGE_PAIR_DELETE_BUTTON_STYLE = "QToolButton { padding: 2px; }"
LANGUAGE_PAIR_DELETE_ICON_SIZE = (16, 16)


# Настройки всплывающего оверлея загрузки в модуле ставок
RATE_TAB_LOADING_OVERLAY_STYLE = "background-color: rgba(0, 0, 0, 204);"
RATE_TAB_LOADING_OVERLAY_LABEL_STYLE = "color: white; font-size: 16px;"
RATE_TAB_LOADING_OVERLAY_SIZE = (200, 100)

# Отступы и размеры списков языков
RATE_TAB_LANG_SECTION_MARGINS = ZERO_MARGINS
RATE_TAB_LANG_SECTION_SPACING = 6
RATE_TAB_TARGET_LAYOUT_MARGINS = ZERO_MARGINS
RATE_TAB_TARGET_LAYOUT_SPACING = 12
RATE_TAB_AVAILABLE_LAYOUT_MARGINS = ZERO_MARGINS
RATE_TAB_AVAILABLE_LAYOUT_SPACING = 6
RATE_TAB_SELECTED_LAYOUT_MARGINS = ZERO_MARGINS
RATE_TAB_SELECTED_LAYOUT_SPACING = 6
RATE_TAB_SELECT_BUTTONS_LAYOUT_MARGINS = ZERO_MARGINS
RATE_TAB_SELECT_BUTTONS_LAYOUT_SPACING = 6
RATE_TAB_LANG_LIST_WIDTH = 260
RATE_TAB_LANG_LIST_HEIGHT = 280
RATE_TAB_SELECTED_LANGUAGES_DISPLAY_STYLE = "border: 1px solid #ccc; padding: 5px;"
RATE_TAB_DELEGATE_PADDING = (5, 5)
RATE_TAB_MISSING_RATE_COLOR = "#FFF3CD"
RATE_TAB_MINIMUM_SIZE = (420, 600)


# Интервал между блоками на правой панели вкладок (основной QVBoxLayout)
RIGHT_PANEL_MAIN_SPACING = 12


# Интервал между блоками на левой панели вкладок (основной QVBoxLayout)
LEFT_PANEL_MAIN_SPACING = 12


# Интервал между строками внутри группы «Информация о проекте» слева
LEFT_PANEL_PROJECT_SECTION_SPACING = 8


# Интервал между элементами внутри группы «Языковые пары» слева
LEFT_PANEL_PAIRS_SECTION_SPACING = 8


# Фиксированная ширина слайдера «Названия языков» на левой панели (в пикселях)
LEFT_PANEL_LANG_MODE_SLIDER_WIDTH = 70


# Максимальная высота поля «Текущие пары» (QTextEdit) на левой панели (в пикселях)
LEFT_PANEL_PAIRS_LIST_MAX_HEIGHT = 100


# Интервал между строками внутри блока «Добавить язык в справочник» слева
LEFT_PANEL_ADD_LANG_SECTION_SPACING = 8


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
# Высота строк таблицы панели ставок (примерно 13 мм при 96 DPI)
RATES_MAPPING_TABLE_ROW_HEIGHT = 54

# Цвет подписи статуса по умолчанию: серый #666666
STATUS_LABEL_DEFAULT_STYLE = "color: #666666;"


# Цвет подписи статуса при успехе: зелёный #107c10
STATUS_LABEL_SUCCESS_STYLE = "color: #107c10;"


# Цвет подписи статуса при ошибке: красный #d13438
STATUS_LABEL_ERROR_STYLE = "color: #d13438;"


# Внутренние отступы ячейки сопоставления языков (SourceTargetCell)
SOURCE_TARGET_CELL_MARGINS = (2, 2, 2, 2)


# Вертикальный зазор между подписью и комбобоксом в ячейке сопоставления
SOURCE_TARGET_CELL_SPACING = 2


# Цвет текста комбобокса выбора названия Excel при необходимости выделения
EXCEL_COMBO_HIGHLIGHT_STYLE = "color: #d97706;"


# Цвет текста кнопки «Импортировать в программу» в отключённом состоянии
IMPORT_BUTTON_ENABLED_STYLE = """
QPushButton {
    background-color: #2563eb;         /* насыщенный синий */
    color: #ffffff;
    border: none;
    border-radius: 6px;
    padding: 8px 18px;
    font-size: 14px;
    font-weight: 600;
}

QPushButton:hover {
    background-color: #1d4ed8;
}

QPushButton:pressed {
    background-color: #1e40af;
}
"""

IMPORT_BUTTON_DISABLED_STYLE = """
QPushButton {
    background-color: #9ca3af;         /* мягкий серый */
    color: #e5e7eb;
    border: none;
    border-radius: 6px;
    padding: 8px 18px;
    font-size: 14px;
    font-weight: 600;
}
"""


# Размеры и отступы диалога импорта расценок
RATES_IMPORT_DIALOG_SIZE = (900, 500)
RATES_IMPORT_DIALOG_MAIN_MARGINS = (15, 15, 15, 15)
RATES_IMPORT_DIALOG_MAIN_SPACING = 12
RATES_IMPORT_DIALOG_SECTION_SPACING = 8
RATES_IMPORT_DIALOG_BROWSE_BUTTON_WIDTH = 70
RATES_IMPORT_DIALOG_CURRENCY_COMBO_WIDTH = 80
RATES_IMPORT_DIALOG_RATE_COMBO_WIDTH = 60
RATES_IMPORT_DIALOG_APPLY_COMBO_WIDTH = 80
RATES_IMPORT_DIALOG_BUTTON_LAYOUT_MARGINS = ZERO_MARGINS
RATES_IMPORT_DIALOG_TABLE_COLUMN_WIDTHS = {
    0: 120,
    1: 120,
    4: 70,
    5: 70,
    6: 70,
}


# Настройки заголовка и палитры приложения
TITLE_BAR_STYLE = "background-color: #ffffff; color: black;"

PALETTE_WINDOW_COLOR = Qt.white
PALETTE_BASE_COLOR = Qt.white
PALETTE_ALTERNATE_BASE_COLOR = QColor(240, 240, 240)
PALETTE_WINDOW_TEXT_COLOR = Qt.black
PALETTE_BUTTON_COLOR = QColor(240, 240, 240)
PALETTE_BUTTON_TEXT_COLOR = Qt.black
PALETTE_TEXT_COLOR = Qt.black
PALETTE_TOOLTIP_BASE_COLOR = Qt.black
PALETTE_TOOLTIP_TEXT_COLOR = Qt.black
PALETTE_BRIGHT_TEXT_COLOR = Qt.red
PALETTE_LINK_COLOR = QColor(42, 130, 218)
PALETTE_HIGHLIGHT_COLOR = QColor(42, 130, 218)
PALETTE_HIGHLIGHTED_TEXT_COLOR = Qt.white
PALETTE_PLACEHOLDER_TEXT_COLOR = QColor("#A0AEC0")
