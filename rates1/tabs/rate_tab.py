
import logging
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
    RATES_ACTION_BUTTON_WIDTH,
    RATES_EXPORT_BUTTON_STYLE,
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
from logic.activity_logger import log_user_action
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
from services.excel_export import apply_excel_styles, table_to_df

# =================== dotenv ===================
from dotenv import load_dotenv

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
        self.source_section_layout = QVBoxLayout()
        self.source_section_layout.setContentsMargins(*RATE_TAB_LANG_SECTION_MARGINS)
        self.source_section_layout.setSpacing(RATE_TAB_LANG_SECTION_SPACING)

        self.source_lang_label = QLabel()
        self.source_section_layout.addWidget(self.source_lang_label)

        self.source_lang_combo = QComboBox()
        self.source_lang_combo.setMinimumWidth(RATE_TAB_LANG_LIST_WIDTH)
        self.source_section_layout.addWidget(self.source_lang_combo)

        self.layout_main.addLayout(self.source_section_layout)

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
        self.toggle_selection_button = QPushButton()
        self.toggle_selection_button.clicked.connect(self.toggle_selection_state)
        self.toggle_selection_button.setStyleSheet(RATE_SELECTION_ACTION_BUTTON_STYLE)
        self.select_buttons_layout.addWidget(self.toggle_selection_button)
        self.select_buttons_layout.addStretch(1)
        self.layout_main.addLayout(self.select_buttons_layout)

        # -- Removed text size slider --

        # --- Выбор ставки и валюты ---
        self.rate_currency_layout = QHBoxLayout()
        self.rate_combo = QComboBox()
        self.currency_combo = QComboBox()
        self.rate_currency_layout.addWidget(self.rate_combo)
        self.rate_currency_layout.addWidget(self.currency_combo)
        self.rate_currency_layout.addStretch(1)
        self.layout_main.addLayout(self.rate_currency_layout)

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
        self.export_button.setStyleSheet(RATES_EXPORT_BUTTON_STYLE)
        self.export_button.setFixedWidth(RATES_ACTION_BUTTON_WIDTH)
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

        self._selected_targets_by_source: Dict[str, List[str]] = {}
        self._active_source: Optional[str] = None
        self._source_order: List[str] = []
        self._export_payloads: Dict[Tuple[str, int], List[Dict[str, object]]] = {}

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
        self.source_lang_combo.currentIndexChanged.connect(
            self._handle_source_combo_change
        )
        self._currency_order = ["USD", "EUR", "RUB", "CNY"]
        self._currency_labels = {
            "USD": "Долл США (USD)",
            "EUR": "Евро (EUR)",
            "RUB": "Рубль (RUB)",
            "CNY": "Юань (CNY)",
        }
        self._rate_labels = {1: "Client rates 1", 2: "Client rates 2"}

        self._update_language_texts()

    # ------------------------------------------------------------------
    # Language helpers
    # ------------------------------------------------------------------
    def _lang(self) -> str:
        return self._lang_getter() if callable(self._lang_getter) else "ru"

    def set_language(self, lang: str) -> None:
        """Update visible texts when the application language changes."""
        self._current_lang = lang
        self._update_language_texts()

    def _update_language_texts(self) -> None:
        lang = self._lang()
        self.load_url_button.setText(tr("Загрузить", lang))
        self.source_lang_label.setText(tr("Исходный язык", lang) + ":")
        self.available_label.setText(tr("Доступные языки", lang) + ":")
        self.available_search.setPlaceholderText(tr("Поиск...", lang))
        self.export_button.setText(tr("Экспорт в Excel", lang))

        self._populate_rate_combo(lang)
        self._populate_currency_combo(lang)
        self._update_selection_summary()

    def set_export_button_width(self, width: int) -> None:
        target = int(width) if isinstance(width, (int, float)) else 0
        hint = self.export_button.sizeHint().width()
        if target <= 0:
            target = hint
        else:
            target = max(target, hint)
        self.export_button.setFixedWidth(target)

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

    def _update_selection_summary(self) -> None:
        lang = self._lang()
        count = self.selected_lang_list.count()
        total = count + self.available_lang_list.count()
        self.selected_label.setText(f"{tr('Выбранные языки', lang)}: {count}")
        if not hasattr(self, "toggle_selection_button"):
            return
        if total <= 0:
            self.toggle_selection_button.setEnabled(False)
            self.toggle_selection_button.setText(tr("Выбрать все", lang))
            return
        self.toggle_selection_button.setEnabled(True)
        if count < total:
            self.toggle_selection_button.setText(tr("Выбрать все", lang))
        else:
            self.toggle_selection_button.setText(tr("Снять выбор", lang))

    # ------------------------------------------------------------------
    # Source language helpers
    # ------------------------------------------------------------------
    def _handle_source_combo_change(self, _index: int) -> None:
        previous_source = self._active_source
        if previous_source:
            previous_targets = [
                self.selected_lang_list.item(i).text()
                for i in range(self.selected_lang_list.count())
            ]
            self._store_targets_for_source(previous_source, previous_targets)

        current = self.source_lang_combo.currentText().strip()
        self.update_target_languages()

    def _store_targets_for_source(
        self, source: str, targets: Optional[List[str]] = None
    ) -> None:
        source = source.strip()
        if not source:
            return
        if targets is None:
            targets = [
                self.selected_lang_list.item(i).text()
                for i in range(self.selected_lang_list.count())
            ]
        unique_targets = []
        seen: Set[str] = set()
        for target in targets:
            normalized = target.strip()
            if not normalized:
                continue
            if normalized in seen:
                continue
            unique_targets.append(normalized)
            seen.add(normalized)

        if unique_targets:
            self._selected_targets_by_source[source] = unique_targets
        else:
            self._selected_targets_by_source.pop(source, None)

    def _sync_active_source_selection(self) -> None:
        current = self._active_source or self.source_lang_combo.currentText().strip()
        if not current:
            return
        self._store_targets_for_source(current)

    def _clear_source_controls(self) -> None:
        self.source_lang_combo.blockSignals(True)
        self.source_lang_combo.clear()
        self.source_lang_combo.blockSignals(False)
        self._active_source = None
        self._source_order = []

    def _populate_source_controls(self, sources: List[str]) -> None:
        self._clear_source_controls()
        if not sources:
            self.available_lang_list.clear()
            self.selected_lang_list.clear()
            self._update_selection_summary()
            return

        self._source_order = list(sources)
        self.source_lang_combo.blockSignals(True)
        self.source_lang_combo.addItems(sources)
        self.source_lang_combo.setCurrentIndex(0)
        self.source_lang_combo.blockSignals(False)

        self._handle_source_combo_change(0)

    def _get_current_selections(self) -> Dict[str, List[str]]:
        selections: Dict[str, List[str]] = {}
        for source in self._source_order:
            targets = self._selected_targets_by_source.get(source, [])
            if targets:
                selections[source] = list(targets)
        return selections

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
            log_user_action(
                "Загрузка ставок MLV_Rates_USD_EUR_RUR_CNY",
                details={"site_id": self.site_id, "file_path": self.file_path},
            )
            token = authenticate_with_msal(self.client_id, self.tenant_id, self.scope)
            if not token:
                log_user_action(
                    "Не получен access_token для MLV_Rates_USD_EUR_RUR_CNY",
                    details={"site_id": self.site_id},
                    level=logging.ERROR,
                )
                return

            df_temp = download_excel_from_sharepoint(token, self.site_id, self.file_path)
            if df_temp is None:
                log_user_action(
                    "Не удалось скачать Excel MLV_Rates_USD_EUR_RUR_CNY",
                    details={"site_id": self.site_id, "file_path": self.file_path},
                    level=logging.ERROR,
                )
                return

            log_user_action(
                "MLV_Rates_USD_EUR_RUR_CNY Excel загружен",
                details={"shape": list(df_temp.shape)},
            )
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
            log_user_action(
                "Загрузка ставок TEP (Source RU)",
                details={"file_id": self.file_id_2, "site_id": self.site_id_2},
            )
            token = authenticate_with_msal(self.client_id, self.tenant_id, self.scope)
            if not token:
                log_user_action(
                    "Не получен access_token для TEP (Source RU)",
                    details={"file_id": self.file_id_2},
                    level=logging.ERROR,
                )
                return

            df_temp = download_excel_by_fileid(
                access_token=token,
                site_id=self.site_id_2,
                file_id=self.file_id_2,
                sheet_name="TEP (Source RU)",
                skiprows=3
            )
            if df_temp is None:
                log_user_action(
                    "Не удалось скачать или прочитать TEP (Source RU)",
                    details={"file_id": self.file_id_2},
                    level=logging.ERROR,
                )
                return

            log_user_action(
                "TEP (Source RU) Excel загружен",
                details={"shape": list(df_temp.shape)},
            )
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
        self._selected_targets_by_source.clear()
        self._active_source = None
        self._source_order = []

        if not self.is_second_file:
            # MLV_Rates_USD_EUR_RUR_CNY
            if self.df.shape[1] < 2:
                return
            source_list = self._clean_language_items(
                self.df.iloc[:, 0].dropna().tolist(),
                {"source"},
            )
            self._populate_source_controls(source_list)
        else:
            # TEP (Source RU)
            if "SourceLang" not in self.df.columns:
                return
            source_list = self._clean_language_items(
                self.df["SourceLang"].dropna().tolist(),
                {"source"},
            )
            self._populate_source_controls(source_list)

        self._apply_auto_selection()

    def update_target_languages(self):
        if self.df is None:
            return
        source_lang = self.source_lang_combo.currentText().strip()

        if self._active_source:
            previous_targets = [
                self.selected_lang_list.item(i).text()
                for i in range(self.selected_lang_list.count())
            ]
            self._store_targets_for_source(self._active_source, previous_targets)

        self.available_lang_list.clear()
        self.selected_lang_list.clear()
        self._update_selection_summary()

        if not source_lang:
            return

        self._active_source = source_lang

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

        saved_targets = self._selected_targets_by_source.get(source_lang, [])
        if saved_targets:
            saved_set = {target.strip() for target in saved_targets}
            for row in range(self.available_lang_list.count() - 1, -1, -1):
                item = self.available_lang_list.item(row)
                if item.text().strip() in saved_set:
                    self.selected_lang_list.addItem(item.text())
                    self.available_lang_list.takeItem(row)
            self.selected_lang_list.sortItems()
        self._update_selection_summary()

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
        self.available_search.clear()
        self._clear_source_controls()
        self._selected_targets_by_source.clear()
        self.available_lang_list.clear()
        self.selected_lang_list.clear()

        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)

        self._update_language_texts()

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
            log_user_action(
                "Исключение при переносе языка в выбранные",
                details={"ошибка": str(e)},
                level=logging.ERROR,
            )

    def move_to_available(self, item):
        try:
            self.available_lang_list.addItem(item.text())
            self.selected_lang_list.takeItem(self.selected_lang_list.row(item))
            self.process_data()
        except Exception as e:
            log_user_action(
                "Исключение при переносе языка в доступные",
                details={"ошибка": str(e)},
                level=logging.ERROR,
            )

    def toggle_selection_state(self) -> None:
        total = self.available_lang_list.count() + self.selected_lang_list.count()
        if total == 0:
            return
        if self.selected_lang_list.count() < total:
            self.select_all_available()
        else:
            self.deselect_all_available()

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
        log_user_action(
            "Запущена обработка ставок",
            details={"второй файл": self.is_second_file},
        )
        self._sync_active_source_selection()
        selections = self._get_current_selections()

        if self.df is None:
            log_user_action(
                "Обработка ставок: нет данных",
                level=logging.ERROR,
            )
            self._update_selection_summary()
            self._emit_current_selection()
            return

        if not selections:
            self._update_selection_summary()
            self.table.setRowCount(0)
            log_user_action(
                "Обработка ставок: не выбраны исходные языки",
                level=logging.ERROR,
            )
            self._emit_current_selection()
            return

        lang = self._lang()
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
        self.table.clearContents()
        self.table.setRowCount(0)
        self.table.viewport().update()

        def _set_item_text(row_index: int, column: int, text: str, *, center: bool = False) -> None:
            item = self.table.item(row_index, column)
            if item is None:
                item = QTableWidgetItem(str(text))
                if center:
                    item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row_index, column, item)
            else:
                item.setText(str(text))
                if center:
                    item.setTextAlignment(Qt.AlignCenter)

        def update_rate_cells(row_index: int, basic, complex_, hour_) -> None:
            values = []
            for value in (basic, complex_, hour_):
                if value is None or (isinstance(value, str) and value == "N/A"):
                    values.append("N/A")
                else:
                    values.append(format_value(value))
            for column, value in zip((2, 3, 4), values):
                _set_item_text(row_index, column, value, center=True)

        if not self.is_second_file:
            log_user_action(
                "Обработка ставок MLV_Rates_USD_EUR_RUR_CNY",
                details={"пары": selections},
            )

            pair_rows: Dict[Tuple[str, str], int] = {}

            def ensure_row(source_text: str, target_text: str) -> int:
                key = (source_text, target_text)
                row_index = pair_rows.get(key)
                if row_index is not None:
                    return row_index

                row_index = self.table.rowCount()
                self.table.insertRow(row_index)
                _set_item_text(row_index, 0, source_text)
                _set_item_text(row_index, 1, target_text)
                for column in (2, 3, 4):
                    _set_item_text(row_index, column, "", center=True)
                pair_rows[key] = row_index
                return row_index

            for source_lang, target_languages in selections.items():
                filtered_df = self.df[self.df.iloc[:, 0] == source_lang]
                source_text = str(source_lang).strip()

                for targ in target_languages:
                    target_text = str(targ).strip()
                    row_found = False

                    for _, row_ in filtered_df.iterrows():
                        if str(row_.iloc[1]).strip() != target_text:
                            continue

                        row_found = True
                        try:
                            basic, complex_, hour_ = self._extract_mlv_rates(
                                row_, selected_currency, rate_number
                            )
                        except Exception as exc:
                            log_user_action(
                                "Ошибка обработки строки MLV_Rates_USD_EUR_RUR_CNY",
                                details={"ошибка": str(exc), "source": source_text, "target": target_text},
                                level=logging.ERROR,
                            )
                            basic = complex_ = hour_ = None

                        row_index = ensure_row(source_text, target_text)
                        update_rate_cells(row_index, basic, complex_, hour_)

                    if not row_found:
                        row_index = ensure_row(source_text, target_text)
                        update_rate_cells(row_index, None, None, None)

        else:
            log_user_action(
                "Обработка ставок TEP (Source RU)",
                details={"пары": selections},
            )
            if "SourceLang" not in self.df.columns or "TargetLang" not in self.df.columns:
                log_user_action(
                    "В таблице TEP отсутствуют необходимые столбцы",
                    level=logging.ERROR,
                )
                return

            pair_rows: Dict[Tuple[str, str], int] = {}

            def ensure_row(source_text: str, target_text: str) -> int:
                key = (source_text, target_text)
                row_index = pair_rows.get(key)
                if row_index is not None:
                    return row_index

                row_index = self.table.rowCount()
                self.table.insertRow(row_index)
                _set_item_text(row_index, 0, source_text)
                _set_item_text(row_index, 1, target_text)
                for column in (2, 3, 4):
                    _set_item_text(row_index, column, "", center=True)
                pair_rows[key] = row_index
                return row_index

            for source_lang, target_languages in selections.items():
                filtered_df = self.df[self.df["SourceLang"] == source_lang]
                source_text = str(source_lang).strip()

                for targ in target_languages:
                    target_text = str(targ).strip()
                    row_found = False

                    for _, row_ in filtered_df.iterrows():
                        if pd.isna(row_["TargetLang"]):
                            continue
                        if str(row_["TargetLang"]).strip() != target_text:
                            continue

                        row_found = True
                        basic, complex_, hour_ = self._extract_tep_rates(
                            row_, selected_currency, rate_number
                        )
                        row_index = ensure_row(source_text, target_text)
                        update_rate_cells(row_index, basic, complex_, hour_)

                    if not row_found:
                        row_index = ensure_row(source_text, target_text)
                        update_rate_cells(row_index, None, None, None)

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
        self._emit_current_selection(selected_currency, rate_number)
        log_user_action(
            "Обработка ставок завершена",
            details={"валюта": selected_currency, "ставка": rate_number},
        )

    def _emit_current_selection(self, selected_currency=None, rate_number=None):
        """Emit the currently displayed rates for embedding into other UIs."""
        if selected_currency is None:
            selected_currency = self.currency_combo.currentData() or "USD"
        if rate_number is None:
            rate_number = self.rate_combo.currentData() or 1

        rows_map: Dict[Tuple[str, str], Dict[str, object]] = {}
        for row in range(self.table.rowCount()):
            entry = {
                "source": self._safe_text(self.table.item(row, 0)),
                "target": self._safe_text(self.table.item(row, 1)),
                "basic": self._parse_float(self._safe_text(self.table.item(row, 2))),
                "complex": self._parse_float(self._safe_text(self.table.item(row, 3))),
                "hour": self._parse_float(self._safe_text(self.table.item(row, 4))),
            }
            key = (entry["source"], entry["target"])
            rows_map.pop(key, None)
            rows_map[key] = entry
        rows = list(rows_map.values())
        self._export_payloads[(selected_currency, rate_number)] = rows

        selections = self._get_current_selections()
        payload = {
            "rows": rows,
            "currency": selected_currency,
            "rate_number": rate_number,
            "rate_type": f"R{rate_number}",
            "source_label": "TEP (Source RU)" if self.is_second_file else "MLV_Rates_USD_EUR_RUR_CNY",
            "is_second_file": self.is_second_file,
            "source_language": ", ".join(selections.keys()),
            "source_languages": list(selections.keys()),
            "targets_by_source": {
                source: list(targets) for source, targets in selections.items()
            },
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

    # ------------------------------------------------------------------
    # Экспорт ставок в Excel (как отображено в таблице GUI)
    # ------------------------------------------------------------------
    def export_rates_to_excel(self):
        if self.table.columnCount() == 0:
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

        current_currency = self.currency_combo.currentData() or "USD"
        current_rate = self.rate_combo.currentData() or 1

        sheet_name = f"R{current_rate}_{current_currency}"
        dataframe_display = table_to_df(self.table)
        numeric_columns = list(dataframe_display.columns[2:])

        dataframe_to_write = dataframe_display.copy()
        na_positions: Dict[str, List[int]] = {}

        for column in numeric_columns:
            values: List[object] = []
            na_rows: List[int] = []
            for row_index, value in enumerate(dataframe_display[column]):
                text = str(value).strip()
                if text.upper() == "N/A":
                    values.append(None)
                    na_rows.append(row_index)
                    continue
                if text == "":
                    values.append(None)
                    continue
                try:
                    values.append(float(text))
                except ValueError:
                    values.append(value)
            dataframe_to_write[column] = values
            na_positions[column] = na_rows

        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            dataframe_to_write.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]

            apply_excel_styles(worksheet, dataframe_to_write, numeric_columns)

            for column_index, column_name in enumerate(dataframe_display.columns, start=1):
                na_rows = na_positions.get(column_name, [])
                for row_index in na_rows:
                    cell = worksheet.cell(row=row_index + 2, column=column_index)
                    cell.value = "N/A"

    def _localized_columns(self, lang: str) -> Tuple[str, str, List[str]]:
        source_col = tr("Исходный язык", lang)
        target_col = tr("Язык перевода", lang)
        columns = [source_col, target_col, "Basic", "Complex", "Hour"]
        return source_col, target_col, columns

    def _empty_rates_dataframe(self, lang: str) -> pd.DataFrame:
        _, _, columns = self._localized_columns(lang)
        return pd.DataFrame(columns=columns)

    def _dataframe_from_rows(
        self, rows: Iterable[Dict[str, object]], lang: str
    ) -> pd.DataFrame:
        source_col, target_col, columns = self._localized_columns(lang)
        data = []
        for entry in rows:
            data.append(
                {
                    source_col: str(entry.get("source", "")),
                    target_col: str(entry.get("target", "")),
                    "Basic": entry.get("basic"),
                    "Complex": entry.get("complex"),
                    "Hour": entry.get("hour"),
                }
            )
        return pd.DataFrame(data, columns=columns)

    def build_rates_dataframe(
        self, source_lang, targets, rate_number, currency, *, lang: str
    ):
        source_col, target_col, columns = self._localized_columns(lang)
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
                    source_col: source_lang,
                    target_col: targ,
                    "Basic": basic,
                    "Complex": complex_,
                    "Hour": hour_
                })
        else:
            if "SourceLang" not in self.df.columns or "TargetLang" not in self.df.columns:
                return self._empty_rates_dataframe(lang)
            filtered_df = self.df[self.df["SourceLang"] == source_lang]
            for targ in targets:
                row_series = filtered_df[filtered_df["TargetLang"].astype(str) == str(targ)]
                if not row_series.empty:
                    basic, complex_, hour_ = self._extract_tep_rates(row_series.iloc[0], currency, rate_number)
                else:
                    basic = complex_ = hour_ = None
                data.append({
                    source_col: source_lang,
                    target_col: targ,
                    "Basic": basic,
                    "Complex": complex_,
                    "Hour": hour_
                })

        return pd.DataFrame(data, columns=columns)

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
