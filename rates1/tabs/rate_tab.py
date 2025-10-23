
import os

# =================== PySide6 ===================
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QLabel,
    QHBoxLayout, QComboBox, QTableWidget, QTableWidgetItem,
    QListWidget, QAbstractItemView,
    QStyledItemDelegate, QHeaderView, QSizePolicy, QLineEdit,
    QDialog, QApplication, QFileDialog
)
from PySide6.QtCore import Qt, QRect, Signal
from PySide6.QtGui import QColor, QFont, QKeySequence, QShortcut

from gui.styles import (
    MLV_RATES_BUTTON_STYLE,
    RATE_SELECTION_ACTION_BUTTON_STYLE,
    RATE_TAB_AVAILABLE_LAYOUT_MARGINS,
    RATE_TAB_AVAILABLE_LAYOUT_SPACING,
    RATE_TAB_DELEGATE_PADDING,
    RATE_TAB_LANG_LIST_HEIGHT,
    RATE_TAB_LANG_LIST_WIDTH,
    RATE_TAB_LANG_SECTION_MARGINS,
    RATE_TAB_LANG_SECTION_SPACING,
    RATE_TAB_LOADING_OVERLAY_LABEL_STYLE,
    RATE_TAB_LOADING_OVERLAY_SIZE,
    RATE_TAB_LOADING_OVERLAY_STYLE,
    RATE_TAB_MISSING_RATE_COLOR,
    RATE_TAB_MINIMUM_SIZE,
    RATE_TAB_SELECT_BUTTONS_LAYOUT_MARGINS,
    RATE_TAB_SELECT_BUTTONS_LAYOUT_SPACING,
    RATE_TAB_SELECTED_LAYOUT_MARGINS,
    RATE_TAB_SELECTED_LAYOUT_SPACING,
    RATE_TAB_TARGET_LAYOUT_MARGINS,
    RATE_TAB_TARGET_LAYOUT_SPACING,
)

# =================== pandas ===================
import pandas as pd

# =================== typing ===================
from typing import Dict, Iterable, List, Optional, Set, Tuple

# =================== rates importer helpers ===================
from logic import rates_importer
from logic.xml_parser_common import language_identity
from logic.translation_config import tr

# =================== Сервисы MS Graph ===================
from services.ms_graph import (
    authenticate_with_msal,
    download_excel_from_sharepoint,
    download_excel_by_fileid
)
from services.excel_export import export_rate_tables

# =================== dotenv ===================
from dotenv import load_dotenv
from utils.history import load_history, add_entry

load_dotenv()

# --------------------------------------------------------------
# Функция для удаления ".0" при выводе
# --------------------------------------------------------------
def format_value(val):
    """
    Если val = "N/A", возвращаем как есть.
    Если val - число, переводим в str. Если оно оканчивается на ".0", убираем.
    """
    if val == "N/A":
        return val
    if isinstance(val, (int, float)):
        s = str(val)
        if s.endswith(".0"):
            s = s[:-2]
        return s
    return str(val)


def safe_float(val):
    """Превращает значение в float, если возможно.

    Возвращает None, если передан нечисловой текст или значение нельзя
    преобразовать к числу. Это позволяет корректно обрабатывать ячейки,
    содержащие ссылки или другие нечисловые данные, не выбрасывая
    исключений при последующем округлении."""
    try:
        val = float(val)
    except (TypeError, ValueError):
        return None
    return None if pd.isna(val) else val

# --------------------------------------------------------------
# Полупрозрачное окошко "Loading..."
# --------------------------------------------------------------
class LoadingOverlay(QDialog):
    def __init__(self, parent=None, text="Loading..."):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setWindowModality(Qt.ApplicationModal)
        self.setStyleSheet(RATE_TAB_LOADING_OVERLAY_STYLE)

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)
        self.setLayout(layout)

        self.label = QLabel(text)
        self.label.setStyleSheet(RATE_TAB_LOADING_OVERLAY_LABEL_STYLE)
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        self.resize(*RATE_TAB_LOADING_OVERLAY_SIZE)
        self.center_on_screen()

    def showEvent(self, event):
        super().showEvent(event)
        self.center_on_screen()

    def center_on_screen(self):
        screen = QApplication.primaryScreen()
        if screen:
            geometry = screen.availableGeometry()
            x = geometry.x() + (geometry.width() - self.width()) // 2
            y = geometry.y() + (geometry.height() - self.height()) // 2
            self.move(x, y)

# --------------------------------------------------------------
# Делегат для отступов
# --------------------------------------------------------------
class CustomDelegate(QStyledItemDelegate):
    def __init__(
        self,
        padding_left=RATE_TAB_DELEGATE_PADDING[0],
        padding_right=RATE_TAB_DELEGATE_PADDING[1],
        parent=None,
    ):
        super().__init__(parent)
        self.padding_left = padding_left
        self.padding_right = padding_right

    def paint(self, painter, option, index):
        if index.column() in [2, 3, 4]:
            option.rect = QRect(
                option.rect.left() + self.padding_left,
                option.rect.top(),
                option.rect.width() - self.padding_left - self.padding_right,
                option.rect.height()
            )
        super().paint(painter, option, index)

class RateTab(QWidget):
    """Rates tab responsible for downloading and previewing rate tables."""

    rates_updated = Signal(object)

    def __init__(self, lang_getter, parent=None):
        super().__init__(parent)
        self._lang_getter = lang_getter
        self._current_lang = self._lang()
        self.setMinimumSize(*RATE_TAB_MINIMUM_SIZE)
        self.layout_main = QVBoxLayout()
        self.setLayout(self.layout_main)

        # --- Кнопки загрузки ---
        self.load_layout = QHBoxLayout()

        self.load_url_button = QPushButton()
        self.load_url_button.clicked.connect(self.load_url)
        self.load_url_button.setStyleSheet(MLV_RATES_BUTTON_STYLE)
        self.load_layout.addWidget(self.load_url_button)
        self.load_layout.addStretch()
        self.load_url_button_2 = None

        self.layout_main.addLayout(self.load_layout)

        # --- Поля ввода языков ---
        self.lang_layout = QHBoxLayout()
        self.source_lang_label = QLabel()
        self.source_lang_combo = QComboBox()
        self.selected_target_lang_label = QLabel()

        self.lang_layout.addWidget(self.source_lang_label)
        self.lang_layout.addWidget(self.source_lang_combo)
        self.lang_layout.addStretch()
        self.lang_layout.addWidget(self.selected_target_lang_label)
        self.layout_main.addLayout(self.lang_layout)

        # --- Списки доступных/выбранных языков ---
        lang_list_width = RATE_TAB_LANG_LIST_WIDTH
        lang_list_height = RATE_TAB_LANG_LIST_HEIGHT

        self.languages_section_layout = QVBoxLayout()
        self.languages_section_layout.setContentsMargins(*RATE_TAB_LANG_SECTION_MARGINS)
        self.languages_section_layout.setSpacing(RATE_TAB_LANG_SECTION_SPACING)

        self.available_search = QLineEdit()
        self.available_search.setClearButtonEnabled(True)
        self.available_search.setMinimumWidth(lang_list_width)
        self.available_search.setMaximumWidth(lang_list_width)
        self.available_search.textChanged.connect(self.filter_available_languages)
        self.languages_section_layout.addWidget(
            self.available_search, alignment=Qt.AlignLeft
        )

        self.target_layout = QHBoxLayout()
        self.target_layout.setContentsMargins(*RATE_TAB_TARGET_LAYOUT_MARGINS)
        self.target_layout.setSpacing(RATE_TAB_TARGET_LAYOUT_SPACING)

        self.available_layout = QVBoxLayout()
        self.available_layout.setContentsMargins(*RATE_TAB_AVAILABLE_LAYOUT_MARGINS)
        self.available_layout.setSpacing(RATE_TAB_AVAILABLE_LAYOUT_SPACING)
        self.available_container = QWidget()
        self.available_container.setLayout(self.available_layout)
        self.available_container.setFixedWidth(lang_list_width)
        self.available_container.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        self.available_label = QLabel()
        self.available_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.available_lang_list = QListWidget()
        self.available_lang_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.available_lang_list.itemDoubleClicked.connect(self.move_to_selected)
        self.available_lang_list.setMinimumWidth(lang_list_width)
        self.available_lang_list.setMaximumWidth(lang_list_width)
        self.available_lang_list.setMinimumHeight(lang_list_height)
        self.available_lang_list.setMaximumHeight(lang_list_height)

        self.available_layout.addWidget(self.available_label)
        self.available_layout.addWidget(self.available_lang_list)

        self.target_layout.addWidget(self.available_container)

        self.selected_layout = QVBoxLayout()
        self.selected_layout.setContentsMargins(*RATE_TAB_SELECTED_LAYOUT_MARGINS)
        self.selected_layout.setSpacing(RATE_TAB_SELECTED_LAYOUT_SPACING)
        self.selected_label = QLabel()
        self.selected_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.selected_lang_list = QListWidget()
        self.selected_lang_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.selected_lang_list.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.selected_lang_list.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.selected_lang_list.itemDoubleClicked.connect(self.move_to_available)
        self.selected_layout.addWidget(self.selected_label)
        self.selected_lang_list.setSortingEnabled(True)
        self.selected_lang_list.setMinimumWidth(lang_list_width)
        self.selected_lang_list.setMaximumWidth(lang_list_width)
        self.selected_lang_list.setMinimumHeight(lang_list_height)
        self.selected_lang_list.setMaximumHeight(lang_list_height)
        self.selected_layout.addWidget(self.selected_lang_list)
        self.selected_container = QWidget()
        self.selected_container.setLayout(self.selected_layout)
        self.selected_container.setFixedWidth(lang_list_width)
        self.selected_container.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        self.target_layout.addWidget(self.selected_container)

        self.languages_section_layout.addLayout(self.target_layout)
        self.layout_main.addLayout(self.languages_section_layout)

        self.select_buttons_layout = QHBoxLayout()
        self.select_buttons_layout.setContentsMargins(*RATE_TAB_SELECT_BUTTONS_LAYOUT_MARGINS)
        self.select_buttons_layout.setSpacing(RATE_TAB_SELECT_BUTTONS_LAYOUT_SPACING)
        self.select_all_button = QPushButton()
        self.deselect_all_button = QPushButton()
        self.select_all_button.clicked.connect(self.select_all_available)
        self.deselect_all_button.clicked.connect(self.deselect_all_available)
        self.select_all_button.setStyleSheet(RATE_SELECTION_ACTION_BUTTON_STYLE)
        self.deselect_all_button.setStyleSheet(RATE_SELECTION_ACTION_BUTTON_STYLE)
        self.select_buttons_layout.addWidget(self.select_all_button)
        self.select_buttons_layout.addWidget(self.deselect_all_button)
        self.select_buttons_layout.addStretch(1)
        self.layout_main.addLayout(self.select_buttons_layout)

        # -- Removed text size slider --

        # --- Выбор ставки (Client rates 1/2) ---
        self.rate_layout = QHBoxLayout()
        self.rate_label = QLabel()
        self.rate_combo = QComboBox()
        self.rate_layout.addWidget(self.rate_label)
        self.rate_layout.addWidget(self.rate_combo)
        self.layout_main.addLayout(self.rate_layout)

        # --- Выбор валюты (USD, EUR, RUB, CNY) ---
        self.currency_layout = QHBoxLayout()
        self.currency_label = QLabel()
        self.currency_combo = QComboBox()
        self.currency_layout.addWidget(self.currency_label)
        self.currency_layout.addWidget(self.currency_combo)
        self.layout_main.addLayout(self.currency_layout)

        # --- История выбора языков ---
        self.history_layout = QHBoxLayout()
        self.history_label = QLabel()
        self.history_combo = QComboBox()
        self.history_layout.addWidget(self.history_label)
        self.history_layout.addWidget(self.history_combo)
        self.layout_main.addLayout(self.history_layout)

        # --- Таблица ---
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.setWordWrap(True)
        self.layout_main.addWidget(self.table)

        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.table)
        self.copy_shortcut.activated.connect(self.copy_to_clipboard)

        # --- Кнопка экспорта ---
        self.export_button = QPushButton()
        self.export_button.clicked.connect(self.export_rates_to_excel)
        self.layout_main.addWidget(self.export_button)

        self.delegate = CustomDelegate(
            padding_left=RATE_TAB_DELEGATE_PADDING[0],
            padding_right=RATE_TAB_DELEGATE_PADDING[1],
            parent=self.table,
        )
        for col in [2, 3, 4]:
            self.table.setItemDelegateForColumn(col, self.delegate)

        self.df = None
        self.default_font = QFont()
        self.setFont(self.default_font)

        self.is_second_file = False
        self._gui_pairs: List[Tuple[str, str]] = []
        self._auto_selection_done = False
        self._excel_matches: List[rates_importer.PairMatch] = []
        self._missing_rate_color = QColor(RATE_TAB_MISSING_RATE_COLOR)

        self.history_data = []
        self.last_saved_selection = None

        # MS Graph настройки через .env
        self.client_id = os.getenv('CLIENT_ID')
        self.tenant_id = os.getenv('TENANT_ID')
        scope_val = os.getenv('SCOPE')
        if scope_val is not None:
            self.scope = [scope_val]
        else:
            self.scope = []
        self.site_id = os.getenv('SITE_ID_1')
        self.file_path = os.getenv('FILE_PATH_1')
        self.site_id_2 = os.getenv('SITE_ID_2')
        self.file_id_2 = os.getenv('FILE_ID_2')

        # Автопересчёт при смене ставки/валюты
        self.rate_combo.currentIndexChanged.connect(self.process_data)
        self.currency_combo.currentIndexChanged.connect(self.process_data)
        self.source_lang_combo.currentIndexChanged.connect(self.update_target_languages)

        # При смене SourceLang - обновляем таргеты
        self.source_lang_combo.currentIndexChanged.connect(self.update_target_languages)

        self.history_combo.currentIndexChanged.connect(self.apply_history_selection)
        self._currency_order = ["USD", "EUR", "RUB", "CNY"]
        self._currency_labels = {
            "USD": "Долл США (USD)",
            "EUR": "Евро (EUR)",
            "RUB": "Рубль (RUB)",
            "CNY": "Юань (CNY)",
        }
        self._rate_labels = {1: "Client rates 1", 2: "Client rates 2"}

        self._update_language_texts()
        self.load_history_combo()

    # ------------------------------------------------------------------
    # Language helpers
    # ------------------------------------------------------------------
    def _lang(self) -> str:
        return self._lang_getter() if callable(self._lang_getter) else "ru"

    def set_language(self, lang: str) -> None:
        """Update visible texts when the application language changes."""
        self._current_lang = lang
        self._update_language_texts()
        self.load_history_combo()

    def _update_language_texts(self) -> None:
        lang = self._lang()
        self.load_url_button.setText(tr("Загрузить", lang))
        self.source_lang_label.setText(tr("Исходный язык", lang) + ":")
        self.selected_label.setText(tr("Выбранные языки", lang) + ":")
        self.available_label.setText(tr("Доступные языки", lang) + ":")
        self.available_search.setPlaceholderText(tr("Поиск...", lang))
        self.select_all_button.setText(tr("Выбрать все", lang))
        self.deselect_all_button.setText(tr("Снять выбор", lang))
        self.rate_label.setText(tr("Выберите ставки", lang) + ":")
        self.currency_label.setText(tr("Выберите валюту", lang) + ":")
        self.history_label.setText(tr("История", lang) + ":")
        self.export_button.setText(tr("Экспорт в Excel", lang))

        self._populate_rate_combo(lang)
        self._populate_currency_combo(lang)
        self._update_selection_summary()
        self._update_history_placeholder(lang)

    def _populate_rate_combo(self, lang: str) -> None:
        current_rate = self.rate_combo.currentData()
        self.rate_combo.blockSignals(True)
        self.rate_combo.clear()
        for rate_id in (1, 2):
            label_key = self._rate_labels[rate_id]
            self.rate_combo.addItem(tr(label_key, lang), userData=rate_id)
        if current_rate not in (1, 2):
            current_rate = 1
        index = self.rate_combo.findData(current_rate)
        self.rate_combo.setCurrentIndex(index if index >= 0 else 0)
        self.rate_combo.blockSignals(False)

    def _populate_currency_combo(self, lang: str) -> None:
        current_code = self.currency_combo.currentData()
        self.currency_combo.blockSignals(True)
        self.currency_combo.clear()
        for code in self._currency_order:
            label_key = self._currency_labels[code]
            self.currency_combo.addItem(tr(label_key, lang), userData=code)
        if current_code not in self._currency_order:
            current_code = self._currency_order[0]
        index = self.currency_combo.findData(current_code)
        self.currency_combo.setCurrentIndex(index if index >= 0 else 0)
        self.currency_combo.blockSignals(False)

    def _update_history_placeholder(self, lang: str) -> None:
        if self.history_combo.count() == 0:
            return
        self.history_combo.setItemText(0, tr("История...", lang))

    def _update_selection_summary(self) -> None:
        lang = self._lang()
        count = self.selected_lang_list.count()
        self.selected_target_lang_label.setText(
            tr("Выбрано языков: {0}", lang).format(count)
        )

    # ----------------------------------------------------------------
    # 1) Загрузка MLV_Rates_USD_EUR_RUR_CNY
    # ----------------------------------------------------------------
    def load_url(self):
        # Disable this button to prevent repeated clicks during loading
        self.load_url_button.setEnabled(False)

        overlay = LoadingOverlay(self, text="Loading MLV_Rates_USD_EUR_RUR_CNY...")
        overlay.show()
        QApplication.processEvents()

        try:
            print("MLV_Rates_USD_EUR_RUR_CNY: site_id =", self.site_id, "file_path =", self.file_path)
            token = authenticate_with_msal(self.client_id, self.tenant_id, self.scope)
            if not token:
                print("Не удалось получить access_token для MLV_Rates_USD_EUR_RUR_CNY")
                return

            df_temp = download_excel_from_sharepoint(token, self.site_id, self.file_path)
            if df_temp is None:
                print("Не удалось скачать MLV_Rates_USD_EUR_RUR_CNY Excel")
                return

            print(f"MLV_Rates_USD_EUR_RUR_CNY Excel загружен: {df_temp.shape}")
            self.df = df_temp
            self.is_second_file = False
            self._auto_selection_done = False
            self.setup_languages()
            self.process_data()
        finally:
            overlay.close()
            self.load_url_button.setEnabled(True)

    # ----------------------------------------------------------------
    # 2) Загрузка TEP (Source RU)
    # ----------------------------------------------------------------
    def load_url_2(self):
        # Disable auxiliary button (if present) to prevent repeated clicks and enable the main one
        if self.load_url_button_2 is not None:
            self.load_url_button_2.setEnabled(False)
        self.load_url_button.setEnabled(True)

        overlay = LoadingOverlay(self, text="Loading TEP (Source RU)...")
        overlay.show()
        QApplication.processEvents()

        try:
            print(f"Начинаем скачивать TEP (Source RU) (fileId): {self.file_id_2}")
            token = authenticate_with_msal(self.client_id, self.tenant_id, self.scope)
            if not token:
                print("Не удалось получить access_token для TEP (Source RU)")
                return

            df_temp = download_excel_by_fileid(
                access_token=token,
                site_id=self.site_id_2,
                file_id=self.file_id_2,
                sheet_name="TEP (Source RU)",
                skiprows=3
            )
            if df_temp is None:
                print("Не удалось скачать/прочитать TEP (Source RU) Excel")
                return

            print(f"TEP (Source RU) Excel загружен, shape = {df_temp.shape}")
            if df_temp.shape[1] > 11:
                df_temp = df_temp.iloc[:, :11]
            rename_map = {}
            if len(df_temp.columns) >= 11:
                rename_map[df_temp.columns[0]] = "SourceLang"
                rename_map[df_temp.columns[1]] = "TargetLang"
                rename_map[df_temp.columns[2]] = "USD_Basic_R1"
                rename_map[df_temp.columns[3]] = "USD_Complex_R1"
                rename_map[df_temp.columns[4]] = "USD_Hourly_R1"
                rename_map[df_temp.columns[5]] = "USD_Basic_R2"
                rename_map[df_temp.columns[6]] = "USD_Complex_R2"
                rename_map[df_temp.columns[7]] = "USD_Hourly_R2"
                rename_map[df_temp.columns[8]] = "RUB_Basic_R1"
                rename_map[df_temp.columns[9]] = "RUB_Complex_R1"
                rename_map[df_temp.columns[10]] = "RUB_Hourly_R1"
            df_temp = df_temp.rename(columns=rename_map)

            self.df = df_temp
            self.is_second_file = True
            self._auto_selection_done = False
            self.setup_languages()
            self.process_data()
        finally:
            overlay.close()
            if self.load_url_button_2 is not None:
                self.load_url_button_2.setEnabled(True)

    # ----------------------------------------------------------------
    # Обновляем список SourceLang
    # ----------------------------------------------------------------
    @staticmethod
    def _clean_language_items(values: Iterable, forbidden: Optional[Set[str]] = None) -> List[str]:
        cleaned: List[str] = []
        seen: Set[str] = set()
        forbidden_norm = {item.casefold() for item in (forbidden or set()) if item}

        for raw in values:
            if raw is None:
                continue
            if isinstance(raw, float) and pd.isna(raw):
                continue
            text = str(raw).strip()
            if not text:
                continue
            lowered = text.casefold()
            if forbidden_norm and lowered in forbidden_norm:
                continue
            if lowered in seen:
                continue
            cleaned.append(text)
            seen.add(lowered)

        return cleaned

    def setup_languages(self):
        if self.df is None or self.df.shape[0] == 0:
            return

        self.available_lang_list.clear()
        self.selected_lang_list.clear()
        self._update_selection_summary()

        if not self.is_second_file:
            # MLV_Rates_USD_EUR_RUR_CNY
            if self.df.shape[1] < 2:
                return
            source_list = self._clean_language_items(
                self.df.iloc[:, 0].dropna().tolist(),
                {"source"},
            )
            self.source_lang_combo.clear()
            self.source_lang_combo.addItems(source_list)
        else:
            # TEP (Source RU)
            if "SourceLang" not in self.df.columns:
                return
            source_list = self._clean_language_items(
                self.df["SourceLang"].dropna().tolist(),
                {"source"},
            )
            self.source_lang_combo.clear()
            self.source_lang_combo.addItems(source_list)

        self._apply_auto_selection()

    def update_target_languages(self):
        if self.df is None:
            return
        source_lang = self.source_lang_combo.currentText()

        self.available_lang_list.clear()
        self.selected_lang_list.clear()
        self._update_selection_summary()

        if not self.is_second_file:
            filtered = self.df[self.df.iloc[:, 0] == source_lang]
            targets = self._clean_language_items(
                filtered.iloc[:, 1].dropna().tolist(),
                {"target"},
            )
            self.available_lang_list.addItems(targets)
        else:
            if "SourceLang" not in self.df.columns or "TargetLang" not in self.df.columns:
                return
            filtered = self.df[self.df["SourceLang"] == source_lang]
            targets = self._clean_language_items(
                filtered["TargetLang"].dropna().tolist(),
                {"target"},
            )
            self.available_lang_list.addItems(targets)

    def set_gui_pairs(self, pairs: Iterable[Tuple[str, str]]):
        cleaned: List[Tuple[str, str]] = []
        for pair in pairs:
            if not isinstance(pair, (tuple, list)) or len(pair) != 2:
                continue
            src, tgt = pair
            src_text = str(src).strip()
            tgt_text = str(tgt).strip()
            if src_text and tgt_text:
                cleaned.append((src_text, tgt_text))
        self._gui_pairs = cleaned
        if self.selected_lang_list.count() > 0:
            self._auto_selection_done = True
            return
        self._auto_selection_done = False
        self._apply_auto_selection()

    def reset_state(self) -> None:
        """Clear loaded datasets and selection state."""
        self.df = None
        self.is_second_file = False
        self._gui_pairs = []
        self._excel_matches = []
        self._auto_selection_done = False
        self.last_saved_selection = None

        self.available_search.clear()
        self.source_lang_combo.blockSignals(True)
        self.source_lang_combo.clear()
        self.source_lang_combo.blockSignals(False)
        self.available_lang_list.clear()
        self.selected_lang_list.clear()

        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)

        self.load_history_combo()
        self._update_language_texts()
        if self.history_combo.count() > 0:
            self.history_combo.setCurrentIndex(0)

        self.rates_updated.emit(
            {
                "rows": [],
                "currency": "",
                "rate_number": 1,
                "rate_type": "",
                "source_label": "",
                "is_second_file": False,
                "source_language": "",
            }
        )

    def set_excel_matches(self, matches: Iterable[rates_importer.PairMatch]) -> None:
        self._excel_matches = [
            match
            for match in matches
            if isinstance(match, rates_importer.PairMatch)
            and getattr(match, "excel_source", "").strip()
            and getattr(match, "excel_target", "").strip()
        ]
        if self.selected_lang_list.count() > 0:
            return
        self._auto_selection_done = False
        self._apply_auto_selection()

    def _apply_auto_selection(self) -> None:
        if self._auto_selection_done:
            return
        if self.df is None:
            return
        if self.source_lang_combo.count() == 0:
            return
        if self.selected_lang_list.count() > 0:
            return

        source_map = self._build_source_map()
        if not source_map:
            return

        selection = self._find_auto_selection(source_map)
        if selection is None:
            return

        idx, display_source, target_norms = selection
        current_index = self.source_lang_combo.currentIndex()
        if current_index != idx:
            self.source_lang_combo.blockSignals(True)
            self.source_lang_combo.setCurrentIndex(idx)
            self.source_lang_combo.blockSignals(False)

        self.update_target_languages()

        source_norm = self._normalize_language_name(display_source)
        expanded_norms = self._expand_target_norms(target_norms, source_norm)
        moved = self._move_targets_to_selected(expanded_norms, source_norm)

        self._auto_selection_done = True
        if moved:
            self.process_data()

    def _build_source_map(self) -> Dict[str, Tuple[int, str]]:
        mapping: Dict[str, Tuple[int, str]] = {}
        for idx in range(self.source_lang_combo.count()):
            text = self.source_lang_combo.itemText(idx).strip()
            if not text:
                continue
            mapping[self._normalize_language_name(text)] = (idx, text)
        return mapping

    def _find_auto_selection(
        self,
        source_map: Dict[str, Tuple[int, str]],
    ) -> Optional[Tuple[int, str, Set[str]]]:
        excel_selection = self._gather_excel_targets(source_map)
        if excel_selection is not None:
            return excel_selection

        gui_selection = self._gather_gui_targets(source_map)
        if gui_selection is not None:
            return gui_selection

        return None

    def _gather_excel_targets(
        self,
        source_map: Dict[str, Tuple[int, str]],
    ) -> Optional[Tuple[int, str, Set[str]]]:
        if not self._excel_matches:
            return None

        order: List[str] = []
        targets_by_source: Dict[str, Set[str]] = {}
        for match in self._excel_matches:
            norm_src = self._normalize_language_name(match.excel_source)
            norm_tgt = self._normalize_language_name(match.excel_target)
            if not norm_src or not norm_tgt:
                continue
            if norm_src not in targets_by_source:
                targets_by_source[norm_src] = set()
                order.append(norm_src)
            targets_by_source[norm_src].add(norm_tgt)

        for src_norm in order:
            if src_norm not in source_map:
                continue
            targets = {norm for norm in targets_by_source[src_norm] if norm}
            if not targets:
                continue
            idx, display_source = source_map[src_norm]
            return idx, display_source, targets

        return None

    def _gather_gui_targets(
        self,
        source_map: Dict[str, Tuple[int, str]],
    ) -> Optional[Tuple[int, str, Set[str]]]:
        if not self._gui_pairs:
            return None

        order: List[str] = []
        targets_by_source: Dict[str, Set[str]] = {}
        for src, tgt in self._gui_pairs:
            norm_src = self._normalize_language_name(src)
            norm_tgt = self._normalize_language_name(tgt)
            if not norm_src or not norm_tgt:
                continue
            if norm_src not in targets_by_source:
                targets_by_source[norm_src] = set()
                order.append(norm_src)
            targets_by_source[norm_src].add(norm_tgt)

        best: Optional[Tuple[int, str, Set[str]]] = None
        for src_norm in order:
            if src_norm not in source_map:
                continue
            targets = {norm for norm in targets_by_source[src_norm] if norm}
            if not targets:
                continue
            idx, display_source = source_map[src_norm]
            if best is None or len(targets) > len(best[2]):
                best = (idx, display_source, targets)

        return best

    def _expand_target_norms(
        self,
        target_norms: Set[str],
        source_norm: str,
    ) -> Set[str]:
        expanded = {norm for norm in target_norms if norm}
        expanded.discard(source_norm)
        return expanded

    def _move_targets_to_selected(
        self,
        target_norms: Set[str],
        source_norm: str,
    ) -> bool:
        if not target_norms:
            self._update_selection_summary()
            return False

        available_norms = {
            self._normalize_language_name(self.available_lang_list.item(i).text())
            for i in range(self.available_lang_list.count())
        }
        target_norms = {
            norm for norm in target_norms if norm and norm in available_norms
        }
        if not target_norms:
            self._update_selection_summary()
            return False

        source_base = source_norm.split("-", 1)[0] if source_norm else ""
        existing_norms = {
            self._normalize_language_name(self.selected_lang_list.item(i).text())
            for i in range(self.selected_lang_list.count())
        }

        moved = False
        row = self.available_lang_list.count() - 1
        while row >= 0:
            item = self.available_lang_list.item(row)
            norm = self._normalize_language_name(item.text())
            if not norm:
                row -= 1
                continue
            base = norm.split("-", 1)[0]
            if norm in existing_norms or base == source_base:
                row -= 1
                continue

            if norm not in target_norms:
                row -= 1
                continue

            self.selected_lang_list.addItem(item.text())
            self.available_lang_list.takeItem(row)
            existing_norms.add(norm)
            moved = True
            row -= 1

        self._update_selection_summary()
        return moved

    @staticmethod
    def _normalize_language_name(value: str) -> str:
        text = str(value).strip()
        if not text:
            return ""
        language, script, territory = language_identity(text)
        if language:
            parts: List[str] = [language.lower()]
            if script:
                parts.append(script.lower())
            if territory:
                parts.append(territory.lower())
            return "-".join(parts)
        try:
            normalized = rates_importer._normalize_language(text)
        except Exception:
            normalized = ""
        return normalized or text.casefold()


    def move_to_selected(self, item):
        try:
            self.selected_lang_list.addItem(item.text())
            self.available_lang_list.takeItem(self.available_lang_list.row(item))
            self.process_data()
        except Exception as e:
            print("move_to_selected => исключение:", e)

    def move_to_available(self, item):
        try:
            self.available_lang_list.addItem(item.text())
            self.selected_lang_list.takeItem(self.selected_lang_list.row(item))
            self.process_data()
        except Exception as e:
            print("move_to_available => исключение:", e)

    def select_all_available(self):
        while self.available_lang_list.count() > 0:
            it = self.available_lang_list.item(0)
            self.selected_lang_list.addItem(it.text())
            self.available_lang_list.takeItem(0)
        self.process_data()

    def deselect_all_available(self):
        while self.selected_lang_list.count() > 0:
            it = self.selected_lang_list.item(0)
            self.available_lang_list.addItem(it.text())
            self.selected_lang_list.takeItem(0)
        self.process_data()

    def filter_available_languages(self, text):
        for i in range(self.available_lang_list.count()):
            it = self.available_lang_list.item(i)
            it.setHidden(text.lower() not in it.text().lower())


    def copy_to_clipboard(self):
        selected_ranges = self.table.selectedRanges()
        if not selected_ranges:
            return
        copied_data = []
        for sr in selected_ranges:
            top_row = sr.topRow()
            bottom_row = sr.bottomRow()
            left_col = sr.leftColumn()
            right_col = sr.rightColumn()

            headers_list = []
            for col in range(left_col, right_col + 1):
                head_item = self.table.horizontalHeaderItem(col)
                headers_list.append(head_item.text() if head_item else "")
            copied_data.append("\t".join(headers_list))

            for row in range(top_row, bottom_row + 1):
                row_data = []
                for col in range(left_col, right_col + 1):
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item else "")
                copied_data.append("\t".join(row_data))

        clipboard_text = "\n".join(copied_data)
        QApplication.clipboard().setText(clipboard_text)

    def process_data(self):
        print("=> process_data() called.")
        if self.df is None:
            print("process_data: df is None => return")
            self._update_selection_summary()
            self._emit_current_selection()
            return

        source_lang = self.source_lang_combo.currentText()
        target_languages = [
            self.selected_lang_list.item(i).text()
            for i in range(self.selected_lang_list.count())
        ]
        lang = self._lang()
        if not target_languages:
            self._update_selection_summary()
            self.table.setRowCount(0)
            print("process_data: нет target_languages => return")
            self._emit_current_selection()
            return

        self._update_selection_summary()

        selected_currency = self.currency_combo.currentData() or "USD"

        rate_number = self.rate_combo.currentData() or 1

        self.table.setColumnCount(5)
        headers = [
            tr("Исходный язык", lang),
            tr("Язык перевода", lang),
            tr("Basic", lang),
            tr("Complex", lang),
            tr("Hour", lang),
        ]
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(0)

        if not self.is_second_file:
            print("process_data: MLV_Rates_USD_EUR_RUR_CNY логика...")
            filtered_df = self.df[self.df.iloc[:, 0] == source_lang]

            for targ in target_languages:
                row_found = False
                fallback_values = None
                for _, row_ in filtered_df.iterrows():
                    if row_.iloc[1] != targ:
                        continue
                    row_found = True
                    try:
                        values = self._extract_mlv_rates(row_, selected_currency, rate_number)
                    except Exception as e:
                        print("Ошибка при обработке MLV_Rates_USD_EUR_RUR_CNY:", e)
                        values = ("N/A", "N/A", "N/A")

                    if any(val != "N/A" for val in values):
                        self._append_rate_row(source_lang, targ, *values)
                        fallback_values = None
                        break

                    if fallback_values is None:
                        fallback_values = values

                if row_found and fallback_values is not None:
                    self._append_rate_row(source_lang, targ, *fallback_values)
                if not row_found:
                    self._append_rate_row(source_lang, targ, "N/A", "N/A", "N/A")

        else:
            print("process_data: TEP (Source RU) логика...")
            if "SourceLang" not in self.df.columns or "TargetLang" not in self.df.columns:
                print("НЕТ 'SourceLang'/'TargetLang' => return")
                return

            filtered_df = self.df[self.df["SourceLang"] == source_lang]
            for targ in target_languages:
                row_found = False
                fallback_values = None
                for _, row_ in filtered_df.iterrows():
                    if not (pd.notnull(row_["TargetLang"]) and str(row_["TargetLang"]) == str(targ)):
                        continue
                    row_found = True
                    values = self._extract_tep_rates(row_, selected_currency, rate_number)

                    if any(val != "N/A" for val in values):
                        self._append_rate_row(source_lang, targ, *values)
                        fallback_values = None
                        break

                    if fallback_values is None:
                        fallback_values = values

                if row_found and fallback_values is not None:
                    self._append_rate_row(source_lang, targ, *fallback_values)
                if not row_found:
                    self._append_rate_row(source_lang, targ, "N/A", "N/A", "N/A")

        self._refresh_missing_rate_highlights()

        # Настройка ширин
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        for col_ in [2, 3, 4]:
            header.setSectionResizeMode(col_, QHeaderView.Stretch)

        for col_ in range(self.table.columnCount()):
            head_item = self.table.horizontalHeaderItem(col_)
            if head_item:
                head_item.setTextAlignment(Qt.AlignCenter)

        header.setStretchLastSection(False)
        self.table.viewport().update()
        self.save_current_selection()
        self._emit_current_selection(selected_currency, rate_number)
        print("=> process_data() done.")

    def _append_rate_row(self, source_lang, target_lang, basic, complex_, hour_):
        rindex = self.table.rowCount()
        self.table.insertRow(rindex)
        self.table.setItem(rindex, 0, QTableWidgetItem(str(source_lang)))
        self.table.setItem(rindex, 1, QTableWidgetItem(str(target_lang)))

        for column, value in enumerate((basic, complex_, hour_), start=2):
            item = QTableWidgetItem(str(value))
            item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(rindex, column, item)

    def _extract_mlv_rates(self, row_, selected_currency, rate_number):
        col_base = None
        basic_round = complex_round = hour_round = 0

        if selected_currency in ["USD", "EUR"]:
            if rate_number == 1:
                if selected_currency == "USD":
                    col_base = 2
                else:
                    col_base = 5
                basic_round = complex_round = 3
            else:
                if selected_currency == "USD":
                    col_base = 8
                else:
                    col_base = 11
                basic_round = complex_round = 3
        elif selected_currency in ["RUB", "CNY"]:
            if rate_number == 1:
                if selected_currency == "RUB":
                    col_base = 15
                else:
                    col_base = 21
                basic_round = complex_round = 2
            else:
                if selected_currency == "RUB":
                    col_base = 18
                else:
                    col_base = 24
                basic_round = complex_round = 2

        if col_base is not None and (col_base + 2) < len(row_):
            bval = safe_float(row_.iloc[col_base])
            cval = safe_float(row_.iloc[col_base + 1])
            hval = safe_float(row_.iloc[col_base + 2])
        else:
            bval = cval = hval = None

        basic = (
            format_value(round(bval, basic_round))
            if bval is not None
            else "N/A"
        )
        complex_ = (
            format_value(round(cval, complex_round))
            if cval is not None
            else "N/A"
        )
        hour_ = (
            format_value(int(round(hval, hour_round)))
            if hval is not None
            else "N/A"
        )

        return basic, complex_, hour_

    def _extract_tep_rates(self, row_, selected_currency, rate_number):
        if selected_currency == "USD":
            if rate_number == 1:
                col_b = "USD_Basic_R1"
                col_c = "USD_Complex_R1"
                col_h = "USD_Hourly_R1"
                br, cr, hr = 3, 3, 0
            else:
                col_b = "USD_Basic_R2"
                col_c = "USD_Complex_R2"
                col_h = "USD_Hourly_R2"
                br, cr, hr = 3, 3, 0
        elif selected_currency == "RUB":
            if rate_number == 1:
                col_b = "RUB_Basic_R1"
                col_c = "RUB_Complex_R1"
                col_h = "RUB_Hourly_R1"
                br, cr, hr = 2, 2, 0
            else:
                col_b = col_c = col_h = None
                br = cr = hr = 0
        else:
            col_b = col_c = col_h = None
            br = cr = hr = 0

        bval = safe_float(row_.get(col_b)) if col_b else None
        cval = safe_float(row_.get(col_c)) if col_c else None
        hval = safe_float(row_.get(col_h)) if col_h else None

        basic = (
            format_value(round(bval, br))
            if bval is not None
            else "N/A"
        )
        complex_ = (
            format_value(round(cval, cr))
            if cval is not None
            else "N/A"
        )
        hour_ = (
            format_value(int(round(hval, hr)))
            if hval is not None
            else "N/A"
        )

        return basic, complex_, hour_

    def _emit_current_selection(self, selected_currency=None, rate_number=None):
        """Emit the currently displayed rates for embedding into other UIs."""
        if selected_currency is None:
            selected_currency = self.currency_combo.currentData() or "USD"
        if rate_number is None:
            rate_number = self.rate_combo.currentData() or 1

        rows = []
        for row in range(self.table.rowCount()):
            rows.append(
                {
                    "source": self._safe_text(self.table.item(row, 0)),
                    "target": self._safe_text(self.table.item(row, 1)),
                    "basic": self._parse_float(self._safe_text(self.table.item(row, 2))),
                    "complex": self._parse_float(self._safe_text(self.table.item(row, 3))),
                    "hour": self._parse_float(self._safe_text(self.table.item(row, 4))),
                }
            )

        payload = {
            "rows": rows,
            "currency": selected_currency,
            "rate_number": rate_number,
            "rate_type": f"R{rate_number}",
            "source_label": "TEP (Source RU)" if self.is_second_file else "MLV_Rates_USD_EUR_RUR_CNY",
            "is_second_file": self.is_second_file,
            "source_language": self.source_lang_combo.currentText(),
        }
        self.rates_updated.emit(payload)

    def _refresh_missing_rate_highlights(self) -> None:
        for row in range(self.table.rowCount()):
            self._apply_missing_rate_highlight(row)

    def _apply_missing_rate_highlight(self, row: int) -> None:
        missing = False
        for col in range(2, min(5, self.table.columnCount())):
            item = self.table.item(row, col)
            text = item.text() if item else ""
            if not text or text == "N/A":
                missing = True
                break

        for col in range(self.table.columnCount()):
            item = self.table.item(row, col)
            if not item:
                continue
            if missing:
                item.setData(Qt.BackgroundRole, self._missing_rate_color)
            else:
                item.setData(Qt.BackgroundRole, None)

    @staticmethod
    def _safe_text(item):
        return "" if item is None else item.text()

    @staticmethod
    def _parse_float(value):
        try:
            if value in ("", "N/A"):
                return None
            return float(value)
        except (TypeError, ValueError):
            return None

    def load_history_combo(self):
        self.history_combo.blockSignals(True)
        self.history_combo.clear()
        lang = self._lang()
        self.history_combo.addItem(tr("История...", lang))
        self.history_data = load_history()
        for entry in self.history_data:
            file_text = "TEP" if entry.get('file', 1) == 2 else "MLV"
            txt = f"{entry['source']} -> {', '.join(entry['targets'])} [{file_text}]"
            self.history_combo.addItem(txt)
        self.history_combo.blockSignals(False)

    def apply_history_selection(self, index):
        if index <= 0:
            return
        entry = self.history_data[index - 1]

        entry_file = entry.get('file', 1)
        if self.df is None or ((entry_file == 2) != self.is_second_file):
            if entry_file == 2:
                self.load_url_2()
            else:
                self.load_url()

        self.source_lang_combo.setCurrentText(entry['source'])
        self.update_target_languages()
        targets_set = set(entry['targets'])
        for i in range(self.available_lang_list.count() - 1, -1, -1):
            it = self.available_lang_list.item(i)
            if it.text() in targets_set:
                self.selected_lang_list.addItem(it.text())
                self.available_lang_list.takeItem(i)
        self.history_combo.setCurrentIndex(0)
        self.process_data()

    def save_current_selection(self):
        source = self.source_lang_combo.currentText()
        targets = [self.selected_lang_list.item(i).text() for i in range(self.selected_lang_list.count())]
        key = (source, tuple(targets))
        if targets and key != self.last_saved_selection:
            add_entry(source, targets, self.is_second_file)
            self.last_saved_selection = key
            self.load_history_combo()

    # ------------------------------------------------------------------
    # Экспорт ставок в Excel (все валюты и R1/R2)
    # ------------------------------------------------------------------
    def export_rates_to_excel(self):
        if self.df is None:
            return

        source_lang = self.source_lang_combo.currentText()
        target_languages = [self.selected_lang_list.item(i).text() for i in range(self.selected_lang_list.count())]
        if not target_languages:
            return

        lang = self._lang()
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            tr("Сохранить ставки", lang),
            "",
            "Excel Files (*.xlsx)",
        )
        if not file_path:
            return
        if not file_path.endswith(".xlsx"):
            file_path += ".xlsx"

        sheets = {}
        for rate_number in [1, 2]:
            for currency in ["USD", "EUR", "RUB", "CNY"]:
                df_rates = self.build_rates_dataframe(source_lang, target_languages, rate_number, currency)
                sheet_name = f"R{rate_number}_{currency}"
                sheets[sheet_name] = df_rates

        export_rate_tables(sheets, file_path)

    def build_rates_dataframe(self, source_lang, targets, rate_number, currency):
        data = []
        if not self.is_second_file:
            filtered_df = self.df[self.df.iloc[:, 0] == source_lang]
            for targ in targets:
                row_series = filtered_df[filtered_df.iloc[:, 1] == targ]
                if not row_series.empty:
                    basic, complex_, hour_ = self._extract_mlv_rates(row_series.iloc[0], currency, rate_number)
                else:
                    basic = complex_ = hour_ = None
                data.append({
                    "Исходный язык": source_lang,
                    "Язык перевода": targ,
                    "Basic": basic,
                    "Complex": complex_,
                    "Hour": hour_
                })
        else:
            if "SourceLang" not in self.df.columns or "TargetLang" not in self.df.columns:
                return pd.DataFrame(columns=["Исходный язык", "Язык перевода", "Basic", "Complex", "Hour"])
            filtered_df = self.df[self.df["SourceLang"] == source_lang]
            for targ in targets:
                row_series = filtered_df[filtered_df["TargetLang"].astype(str) == str(targ)]
                if not row_series.empty:
                    basic, complex_, hour_ = self._extract_tep_rates(row_series.iloc[0], currency, rate_number)
                else:
                    basic = complex_ = hour_ = None
                data.append({
                    "Исходный язык": source_lang,
                    "Язык перевода": targ,
                    "Basic": basic,
                    "Complex": complex_,
                    "Hour": hour_
                })

        return pd.DataFrame(data, columns=["Исходный язык", "Язык перевода", "Basic", "Complex", "Hour"])

    def _extract_mlv_rates(self, row_, currency, rate_number):
        col_base = None
        hr = 0
        if currency in ["USD", "EUR"]:
            if rate_number == 1:
                col_base = 2 if currency == "USD" else 5
                br = cr = 3
            else:
                col_base = 8 if currency == "USD" else 11
                br = cr = 3
        elif currency in ["RUB", "CNY"]:
            if rate_number == 1:
                col_base = 15 if currency == "RUB" else 21
                br = cr = 2
            else:
                col_base = 18 if currency == "RUB" else 24
                br = cr = 2
        if col_base is not None and (col_base + 2) < len(row_):
            bval = safe_float(row_.iloc[col_base])
            cval = safe_float(row_.iloc[col_base + 1])
            hval = safe_float(row_.iloc[col_base + 2])
            basic = round(bval, br) if bval is not None else None
            complex_ = round(cval, cr) if cval is not None else None
            hour_ = int(round(hval, hr)) if hval is not None else None
        else:
            basic = complex_ = hour_ = None
        return basic, complex_, hour_

    def _extract_tep_rates(self, row_, currency, rate_number):
        if currency == "USD":
            if rate_number == 1:
                col_b = "USD_Basic_R1"
                col_c = "USD_Complex_R1"
                col_h = "USD_Hourly_R1"
                br = cr = 3
            else:
                col_b = "USD_Basic_R2"
                col_c = "USD_Complex_R2"
                col_h = "USD_Hourly_R2"
                br = cr = 3
            hr = 0
        elif currency == "RUB":
            if rate_number == 1:
                col_b = "RUB_Basic_R1"
                col_c = "RUB_Complex_R1"
                col_h = "RUB_Hourly_R1"
                br = cr = 2
                hr = 0
            else:
                col_b = col_c = col_h = None
                br = cr = hr = 0
        else:
            col_b = col_c = col_h = None
            br = cr = hr = 0

        bval = safe_float(row_.get(col_b)) if col_b else None
        cval = safe_float(row_.get(col_c)) if col_c else None
        hval = safe_float(row_.get(col_h)) if col_h else None

        basic = round(bval, br) if bval is not None else None
        complex_ = round(cval, cr) if cval is not None else None
        hour_ = int(round(hval, hr)) if hval is not None else None
        return basic, complex_, hour_
