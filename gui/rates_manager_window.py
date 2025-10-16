"""Embedded version of the standalone rates utility."""

from __future__ import annotations

import logging
from functools import partial
from typing import Dict, Iterable, List, Optional, Set, Tuple

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QComboBox,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from gui.styles import (
    EXCEL_COMBO_HIGHLIGHT_STYLE,
    IMPORT_BUTTON_DISABLED_STYLE,
    IMPORT_BUTTON_ENABLED_STYLE,
    RATES_IMPORT_DIALOG_STYLE,
    RATES_MAPPING_APPLY_COMBO_WIDTH,
    RATES_MAPPING_CONTROLS_SPACING,
    RATES_MAPPING_LAYOUT_MARGINS,
    RATES_MAPPING_LAYOUT_SPACING,
    RATES_MAPPING_TABLE_COLUMN_WIDTHS,
    RATES_MAPPING_TABLE_ROW_HEIGHT,
    RATES_WINDOW_LAYOUT_MARGINS,
    RATES_WINDOW_LAYOUT_SPACING,
    RATES_WINDOW_SPLITTER_SIZES,
    RATES_WINDOW_SPLITTER_STRETCH_FACTORS,
    SOURCE_TARGET_CELL_MARGINS,
    SOURCE_TARGET_CELL_SPACING,
    STATUS_LABEL_DEFAULT_STYLE,
    STATUS_LABEL_SUCCESS_STYLE,
)
from logic import rates_importer
from logic.translation_config import tr
from rates1 import RateTab

RateRow = Dict[str, Optional[float]]

logger = logging.getLogger(__name__)


class SourceTargetCell(QWidget):
    """Composite cell that shows GUI value and editable Excel value."""

    excel_changed = Signal(str)

    def __init__(self, main_value: str, lang_getter, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._lang_getter = lang_getter
        self._main_value = main_value

        layout = QVBoxLayout(self)
        layout.setContentsMargins(*SOURCE_TARGET_CELL_MARGINS)
        layout.setSpacing(SOURCE_TARGET_CELL_SPACING)

        self.main_label = QLabel(main_value)
        font = self.main_label.font()
        font.setBold(True)
        self.main_label.setFont(font)
        layout.addWidget(self.main_label)

        self.excel_combo = QComboBox()
        self.excel_combo.setEditable(True)
        self.excel_combo.setInsertPolicy(QComboBox.NoInsert)
        self.excel_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.excel_combo.currentTextChanged.connect(self._on_excel_changed)
        self.excel_combo.setContextMenuPolicy(Qt.CustomContextMenu)
        self.excel_combo.customContextMenuRequested.connect(self._show_combo_menu)
        layout.addWidget(self.excel_combo)

        self.main_label.setContextMenuPolicy(Qt.CustomContextMenu)
        self.main_label.customContextMenuRequested.connect(self._show_label_menu)

        self.setContextMenuPolicy(Qt.DefaultContextMenu)
        self.update_highlight()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def set_main_value(self, text: str) -> None:
        self._main_value = text
        self.main_label.setText(text)
        self.update_highlight()

    def set_language_names(self, names: Iterable[str]) -> None:
        current = self.excel_combo.currentText()
        self.excel_combo.blockSignals(True)
        self.excel_combo.clear()
        self.excel_combo.addItems(list(names))
        self.excel_combo.setEditText(current)
        self.excel_combo.blockSignals(False)
        self.update_highlight()

    def set_excel_text(self, text: str, emit: bool = False) -> None:
        current = self.excel_combo.currentText()
        if current == text:
            self.update_highlight()
            if emit:
                self.excel_changed.emit(text)
            return
        self.excel_combo.blockSignals(True)
        self.excel_combo.setEditText(text or "")
        self.excel_combo.blockSignals(False)
        self.update_highlight()
        if emit:
            self.excel_changed.emit(self.excel_combo.currentText())

    def excel_text(self) -> str:
        return self.excel_combo.currentText()

    def main_value(self) -> str:
        return self._main_value

    def update_highlight(self) -> None:
        text = self.excel_combo.currentText().strip()
        highlight = bool(text) and text != self._main_value
        self.excel_combo.setStyleSheet(
            EXCEL_COMBO_HIGHLIGHT_STYLE if highlight else ""
        )

    # ------------------------------------------------------------------
    # Qt events
    # ------------------------------------------------------------------
    def contextMenuEvent(self, event) -> None:  # noqa: D401 - Qt override
        self._show_menu(event.globalPos())

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _lang(self) -> str:
        return self._lang_getter() if callable(self._lang_getter) else "ru"

    def _on_excel_changed(self, _text: str) -> None:
        self.update_highlight()
        self.excel_changed.emit(self.excel_combo.currentText())

    def _show_combo_menu(self, pos) -> None:
        self._show_menu(self.excel_combo.mapToGlobal(pos))

    def _show_label_menu(self, pos) -> None:
        self._show_menu(self.main_label.mapToGlobal(pos))

    def _show_menu(self, global_pos) -> None:
        menu = QMenu(self)
        copy_action = menu.addAction(
            tr("Скопировать основное → Excel", self._lang())
        )
        reset_action = menu.addAction(tr("Сбросить Excel-значение", self._lang()))
        chosen = menu.exec(global_pos)
        if chosen == copy_action:
            self.set_excel_text(self._main_value, emit=True)
        elif chosen == reset_action:
            self.set_excel_text("", emit=True)

class RatesMappingWidget(QWidget):
    """Widget responsible for mapping remote rates to GUI language pairs."""

    import_requested = Signal()

    def __init__(self, lang_getter, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._lang_getter = lang_getter
        self._pairs: List[Tuple[str, str]] = []
        self._rates: rates_importer.RatesMap = {}
        self._name_to_code: Dict[str, str] = {}
        self._lang_names: List[str] = []
        self._source_label: str = ""
        self._currency: str = ""
        self._rate_type: str = ""
        self._last_matches: List[rates_importer.PairMatch] = []
        self._manual_excel_cells: Set[Tuple[int, int]] = set()
        self._updating_excel_from_matches: bool = False

        self._setup_ui()
        self.setStyleSheet(RATES_IMPORT_DIALOG_STYLE)
        self._update_language_texts()
        self._update_status()
        self._update_import_button_state()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(*RATES_MAPPING_LAYOUT_MARGINS)
        layout.setSpacing(RATES_MAPPING_LAYOUT_SPACING)

        apply_layout = QHBoxLayout()
        apply_layout.setSpacing(RATES_MAPPING_CONTROLS_SPACING)

        self.apply_label = QLabel()
        apply_layout.addWidget(self.apply_label)

        self.apply_combo = QComboBox()
        self.apply_combo.setFixedWidth(RATES_MAPPING_APPLY_COMBO_WIDTH)
        apply_layout.addWidget(self.apply_combo)
        apply_layout.addStretch()

        layout.addLayout(apply_layout)

        status_layout = QHBoxLayout()
        status_layout.addStretch()
        self.status_label = QLabel(tr("Данные не загружены", self._lang()))
        self.status_label.setStyleSheet(STATUS_LABEL_DEFAULT_STYLE)
        status_layout.addWidget(self.status_label)
        layout.addLayout(status_layout)

        self.table = QTableWidget(0, 3)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)

        header = self.table.horizontalHeader()
        for column, width in enumerate(RATES_MAPPING_TABLE_COLUMN_WIDTHS):
            header.setSectionResizeMode(column, QHeaderView.Interactive)
            self.table.setColumnWidth(column, width)
        vertical_header = self.table.verticalHeader()
        vertical_header.setDefaultSectionSize(RATES_MAPPING_TABLE_ROW_HEIGHT)
        vertical_header.setMinimumSectionSize(RATES_MAPPING_TABLE_ROW_HEIGHT)
        self.table.itemChanged.connect(self._handle_item_changed)
        layout.addWidget(self.table, 1)

        self._rate_headers = {"basic": "Basic", "complex": "Complex", "hour": "Hour"}
        self._rate_order = ("basic", "complex", "hour")
        self._populate_apply_combo(self._lang())
        self._update_table_headers()
        self._rate_values: List[Dict[str, str]] = []
        self._updating_rate_item = False
        self.apply_combo.currentIndexChanged.connect(self._handle_rate_mode_change)

        import_layout = QHBoxLayout()
        import_layout.addStretch()
        self.import_btn = QPushButton()
        self.import_btn.clicked.connect(self.import_requested.emit)
        import_layout.addWidget(self.import_btn)
        layout.addLayout(import_layout)

    def set_language(self, _lang: str) -> None:
        self._update_language_texts()
        self._update_status()
        self._refresh_all_rate_display()
        self._update_import_button_state()

    def _update_language_texts(self) -> None:
        lang = self._lang()
        self.apply_label.setText(tr("Применить", lang) + ":")
        self.import_btn.setText(tr("Импортировать в программу", lang))
        self._populate_apply_combo(lang)
        self._update_table_headers()

    def _populate_apply_combo(self, lang: str) -> None:
        current_key = self.selected_rate_key()
        self.apply_combo.blockSignals(True)
        self.apply_combo.clear()
        for key in self._rate_order:
            label_key = self._rate_headers.get(key, "")
            self.apply_combo.addItem(tr(label_key, lang), userData=key)
        index = self.apply_combo.findData(current_key)
        self.apply_combo.setCurrentIndex(index if index >= 0 else 0)
        self.apply_combo.blockSignals(False)

    def _update_table_headers(self) -> None:
        lang = self._lang()
        headers = [
            tr("Исходный язык", lang),
            tr("Язык перевода", lang),
            tr(self._rate_headers.get(self.selected_rate_key(), "Basic"), lang),
        ]
        for idx, text in enumerate(headers):
            item = self.table.horizontalHeaderItem(idx)
            if item is None:
                item = QTableWidgetItem(text)
                self.table.setHorizontalHeaderItem(idx, item)
            else:
                item.setText(text)

    def _update_import_button_state(self) -> None:
        enabled = self._can_import()
        self.import_btn.setEnabled(enabled)
        self.import_btn.setStyleSheet(
            IMPORT_BUTTON_ENABLED_STYLE if enabled else IMPORT_BUTTON_DISABLED_STYLE
        )

    def _can_import(self) -> bool:
        if not self._rates or self.table.rowCount() == 0:
            return False
        key = self.selected_rate_key()
        for row in range(self.table.rowCount()):
            src_widget = self.table.cellWidget(row, 0)
            tgt_widget = self.table.cellWidget(row, 1)
            src_text = (
                src_widget.excel_text().strip()
                if isinstance(src_widget, SourceTargetCell)
                else ""
            )
            tgt_text = (
                tgt_widget.excel_text().strip()
                if isinstance(tgt_widget, SourceTargetCell)
                else ""
            )
            storage = self._ensure_rate_storage(row)
            value = storage.get(key, "").strip()
            if not src_text or not tgt_text or not value:
                return False
        return True

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def set_pairs(self, pairs: Iterable[Tuple[str, str]]) -> None:
        self._pairs = list(pairs)
        self._last_matches = []
        self._rebuild_rows()

    def set_rates_data(
        self,
        rates: rates_importer.RatesMap,
        source_label: str,
        currency: str,
        rate_type: str,
    ) -> None:
        self._rates = rates or {}
        self._source_label = source_label
        self._currency = currency
        self._rate_type = rate_type

        lang_codes = set()
        for src, tgt in self._rates:
            lang_codes.add(src)
            lang_codes.add(tgt)

        self._name_to_code = {
            rates_importer._language_name(code): code for code in sorted(lang_codes)
        }
        self._lang_names = sorted(self._name_to_code.keys())

        self._refresh_language_combos()
        self._apply_matches(auto_fill=True)
        self._update_status()

    def auto_fill_from_rates(self, force: bool = False) -> None:
        _ = force  # параметр сохранён для совместимости с существующими вызовами
        for row in range(self.table.rowCount()):
            self._update_rate_from_row(row)
        self._update_import_button_state()

    def selected_rates(self) -> List[rates_importer.PairMatch]:
        matches: List[rates_importer.PairMatch] = []
        for row in range(self.table.rowCount()):
            if row < len(self._pairs):
                gui_src, gui_tgt = self._pairs[row]
            else:
                gui_src, gui_tgt = "", ""

            excel_src = self._lang_cell_excel_text(row, 0)
            excel_tgt = self._lang_cell_excel_text(row, 1)

            storage = self._ensure_rate_storage(row)
            basic = storage.get("basic", "")
            complex_ = storage.get("complex", "")
            hour = storage.get("hour", "")

            rate: Optional[RateRow]
            if basic or complex_ or hour:
                rate = self._parse_rate(basic, complex_, hour)
            else:
                rate = None

            matches.append(
                rates_importer.PairMatch(
                    gui_source=gui_src,
                    gui_target=gui_tgt,
                    excel_source=excel_src,
                    excel_target=excel_tgt,
                    rates=rate,
                )
            )

        return matches

    def selected_rate_key(self) -> str:
        data = self.apply_combo.currentData()
        return data if isinstance(data, str) else "basic"

    def current_currency(self) -> str:
        return self._currency.strip().upper() if self._currency else ""

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _lang(self) -> str:
        return self._lang_getter() if callable(self._lang_getter) else "ru"

    def _rebuild_rows(self) -> None:
        current_rows = self.table.rowCount()
        target_rows = len(self._pairs)
        if target_rows != current_rows:
            self.table.setRowCount(target_rows)

        if len(self._rate_values) < target_rows:
            self._rate_values.extend(self._empty_rate_dict() for _ in range(target_rows - len(self._rate_values)))
        elif len(self._rate_values) > target_rows:
            self._rate_values = self._rate_values[:target_rows]

        self._manual_excel_cells = {
            key
            for key in self._manual_excel_cells
            if key[0] < target_rows and key[1] in (0, 1)
        }

        for row, (src, tgt) in enumerate(self._pairs):
            self.table.setRowHeight(row, RATES_MAPPING_TABLE_ROW_HEIGHT)
            src_cell = self._ensure_lang_cell(row, 0)
            if src_cell.main_value() != src:
                self._manual_excel_cells.discard((row, 0))
            src_cell.set_main_value(src)

            tgt_cell = self._ensure_lang_cell(row, 1)
            if tgt_cell.main_value() != tgt:
                self._manual_excel_cells.discard((row, 1))
            tgt_cell.set_main_value(tgt)

            rate_item = self.table.item(row, 2)
            if rate_item is None:
                rate_item = QTableWidgetItem("")
                rate_item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row, 2, rate_item)

            self._refresh_rate_display(row)

        self._update_import_button_state()

    def _refresh_language_combos(self) -> None:
        for row in range(self.table.rowCount()):
            for column in (0, 1):
                cell = self._ensure_lang_cell(row, column)
                cell.set_language_names(self._lang_names)

    def _apply_matches(self, auto_fill: bool) -> None:
        if not self._pairs:
            return

        matches = rates_importer.match_pairs(self._pairs, self._rates)
        self._last_matches = matches
        for row, match in enumerate(matches):
            src_cell = self._ensure_lang_cell(row, 0)
            tgt_cell = self._ensure_lang_cell(row, 1)
            if not self._is_manual_excel_cell(row, 0):
                self._set_excel_text_from_match(src_cell, row, 0, match.excel_source)
            if not self._is_manual_excel_cell(row, 1):
                self._set_excel_text_from_match(tgt_cell, row, 1, match.excel_target)

            if auto_fill and match.rates:
                self._apply_rate_values(row, match.rates)

        self._update_import_button_state()

    def _ensure_lang_cell(self, row: int, column: int, main_value: Optional[str] = None) -> SourceTargetCell:
        widget = self.table.cellWidget(row, column)
        if isinstance(widget, SourceTargetCell):
            if main_value is not None:
                widget.set_main_value(main_value)
            return widget
        cell = SourceTargetCell(main_value or "", self._lang_getter)
        cell.excel_changed.connect(partial(self._handle_excel_value_change, row, column))
        self.table.setCellWidget(row, column, cell)
        cell.set_language_names(self._lang_names)
        return cell

    def _handle_excel_value_change(self, row: int, column: int, text: str) -> None:
        if not self._updating_excel_from_matches:
            key = (row, column)
            if text.strip():
                self._manual_excel_cells.add(key)
            else:
                self._manual_excel_cells.discard(key)
        self._update_rate_from_row(row)
        self._update_import_button_state()

    def _set_excel_text_from_match(
        self, cell: SourceTargetCell, row: int, column: int, value: Optional[str]
    ) -> None:
        self._updating_excel_from_matches = True
        try:
            cell.set_excel_text(value or "")
        finally:
            self._updating_excel_from_matches = False
        if not (value or "").strip():
            self._manual_excel_cells.discard((row, column))

    def _is_manual_excel_cell(self, row: int, column: int) -> bool:
        return (row, column) in self._manual_excel_cells

    def _handle_rate_mode_change(self, _text: str) -> None:
        self._update_table_headers()
        self._refresh_all_rate_display()
        self._update_import_button_state()

    def _update_rate_from_row(self, row: int) -> None:
        src_cell = self._ensure_lang_cell(row, 0)
        tgt_cell = self._ensure_lang_cell(row, 1)
        src_code = self._name_to_code.get(src_cell.excel_text())
        tgt_code = self._name_to_code.get(tgt_cell.excel_text())
        if not src_code or not tgt_code:
            self._update_import_button_state()
            return

        rate = self._rates.get((src_code, tgt_code))
        if rate:
            self._apply_rate_values(row, rate)
        else:
            self._clear_rate_values(row)
        self._update_import_button_state()

    def matched_pairs(self) -> List[rates_importer.PairMatch]:
        return list(self._last_matches)

    def _apply_rate_values(self, row: int, rate: RateRow) -> None:
        basic = rate.get("basic")
        complex_ = rate.get("complex")
        hour = rate.get("hour")
        self._set_rate_value(row, "basic", basic)
        self._set_rate_value(row, "complex", complex_)
        self._set_rate_value(row, "hour", hour)
        self._refresh_rate_display(row)

    def _clear_rate_values(self, row: int) -> None:
        storage = self._ensure_rate_storage(row)
        for key in ("basic", "complex", "hour"):
            storage[key] = ""
        self._refresh_rate_display(row)
        self._update_import_button_state()

    def _set_rate_value(self, row: int, key: str, value: Optional[float]) -> None:
        storage = self._ensure_rate_storage(row)
        if value is None:
            storage[key] = ""
        else:
            storage[key] = f"{value:.2f}"

    def _ensure_rate_storage(self, row: int) -> Dict[str, str]:
        while len(self._rate_values) <= row:
            self._rate_values.append(self._empty_rate_dict())
        return self._rate_values[row]

    def _refresh_rate_display(self, row: int) -> None:
        if row >= self.table.rowCount():
            return
        item = self.table.item(row, 2)
        if item is None:
            item = QTableWidgetItem("")
            item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 2, item)
        key = self.selected_rate_key()
        storage = self._ensure_rate_storage(row)
        value = storage.get(key, "")
        self._updating_rate_item = True
        item.setText(value)
        self._updating_rate_item = False
        self._update_import_button_state()

    def _refresh_all_rate_display(self) -> None:
        for row in range(self.table.rowCount()):
            self._refresh_rate_display(row)

    def _update_rate_header(self, selected_key: str) -> None:
        self._update_table_headers()

    def _parse_rate(self, basic: str, complex_: str, hour: str) -> Optional[RateRow]:
        def _to_float(value: str) -> Optional[float]:
            try:
                return float(value) if value else None
            except (TypeError, ValueError):
                return None

        b_val = _to_float(basic)
        c_val = _to_float(complex_)
        h_val = _to_float(hour)
        if b_val is None and c_val is None and h_val is None:
            return None
        return {"basic": b_val, "complex": c_val, "hour": h_val}

    def _update_status(self) -> None:
        lang = self._lang()
        if not self._rates:
            self.status_label.setText(tr("Данные не загружены", lang))
            self.status_label.setStyleSheet(STATUS_LABEL_DEFAULT_STYLE)
            self._update_import_button_state()
            return

        pieces = [f"{tr('Загружено пар', lang)}: {len(self._rates)}"]
        if self._source_label:
            pieces.append(f"{tr('Источник', lang)}: {self._source_label}")
        if self._currency:
            pieces.append(f"{tr('Валюта', lang)}: {self._currency}")
        if self._rate_type:
            pieces.append(f"{tr('Тип ставки', lang)}: {self._rate_type}")

        self.status_label.setText(" | ".join(pieces))
        self.status_label.setStyleSheet(STATUS_LABEL_SUCCESS_STYLE)
        self._update_import_button_state()

    def _empty_rate_dict(self) -> Dict[str, str]:
        return {"basic": "", "complex": "", "hour": ""}

    def _handle_item_changed(self, item: QTableWidgetItem) -> None:
        if self._updating_rate_item or item.column() != 2:
            return
        storage = self._ensure_rate_storage(item.row())
        storage[self.selected_rate_key()] = item.text()
        self._update_import_button_state()

    def _lang_cell_excel_text(self, row: int, column: int) -> str:
        widget = self.table.cellWidget(row, column)
        if isinstance(widget, SourceTargetCell):
            return widget.excel_text()
        return ""


class RatesManagerWindow(QMainWindow):
    """Main window embedding the remote rates explorer."""

    def __init__(self, main_window) -> None:
        super().__init__(main_window)
        self._main_window = main_window
        self._current_pairs: List[Tuple[str, str]] = []

        self.setWindowTitle(tr("Панель ставок", main_window.gui_lang))
        self.resize(main_window.size())
        self.move(main_window.pos())

        central = QWidget()
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(*RATES_WINDOW_LAYOUT_MARGINS)
        root_layout.setSpacing(RATES_WINDOW_LAYOUT_SPACING)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        root_layout.addWidget(splitter)

        self.rate_tab = RateTab(lambda: self._main_window.gui_lang)
        splitter.addWidget(self.rate_tab)

        self.mapping_widget = RatesMappingWidget(lambda: self._main_window.gui_lang, self)
        splitter.addWidget(self.mapping_widget)

        splitter.setStretchFactor(0, RATES_WINDOW_SPLITTER_STRETCH_FACTORS[0])
        splitter.setStretchFactor(1, RATES_WINDOW_SPLITTER_STRETCH_FACTORS[1])
        splitter.setSizes(RATES_WINDOW_SPLITTER_SIZES)

        self.setCentralWidget(central)

        self.rate_tab.rates_updated.connect(self._handle_rate_payload)
        self.mapping_widget.import_requested.connect(self._apply_to_main_window)

    def set_language(self, lang: str) -> None:
        """Update visible texts to the requested language."""
        self.setWindowTitle(tr("Панель ставок", lang))
        self.mapping_widget.set_language(lang)
        self.rate_tab.set_language(lang)

    # ------------------------------------------------------------------
    # Public helpers
    # ------------------------------------------------------------------
    def update_pairs(self, pairs: Iterable[Tuple[str, str]]) -> None:
        self._current_pairs = list(pairs)
        self.mapping_widget.set_pairs(self._current_pairs)
        self.rate_tab.set_gui_pairs(self._current_pairs)

    # ------------------------------------------------------------------
    # Slots
    # ------------------------------------------------------------------
    def _handle_rate_payload(self, payload: Dict[str, object]) -> None:
        rates = self._build_rates_map(payload)
        source_label = str(payload.get("source_label", ""))
        currency = str(payload.get("currency", ""))
        rate_type = str(payload.get("rate_type", ""))
        self.mapping_widget.set_rates_data(rates, source_label, currency, rate_type)
        matches = self.mapping_widget.matched_pairs()
        if matches:
            self.rate_tab.set_excel_matches(matches)

    def _apply_to_main_window(self) -> None:
        matches = self.mapping_widget.selected_rates()
        rate_key = self.mapping_widget.selected_rate_key()
        if not matches:
            QMessageBox.information(
                self,
                tr("Ставки", self._main_window.gui_lang),
                tr("Данные не загружены", self._main_window.gui_lang),
            )
            return
        currency = self.mapping_widget.current_currency()
        success = True
        error_message = ""
        try:
            if currency:
                self._main_window.set_currency_code(currency)
            self._main_window._apply_rates_from_matches(matches, rate_key)
        except Exception as exc:  # pragma: no cover - GUI level safeguard
            success = False
            error_message = str(exc)
            logger.exception("Failed to import rates from mapping widget")

        self._show_import_result(success, error_message)

    # ------------------------------------------------------------------
    # Data helpers
    # ------------------------------------------------------------------
    def _show_import_result(self, success: bool, details: str = "") -> None:
        lang = self._main_window.gui_lang
        box = QMessageBox(self)
        box.setWindowTitle(tr("Импорт ставок", lang))
        box.setIcon(QMessageBox.Information if success else QMessageBox.Critical)
        text_key = "Ставки успешно добавлены." if success else "Не удалось импортировать ставки."
        message = tr(text_key, lang)
        if not success and details:
            message += f"\n{details}"
        box.setText(message)
        ok_button = box.addButton(tr("Ок", lang), QMessageBox.AcceptRole)
        close_button = box.addButton(tr("Закрыть панель", lang), QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() == close_button:
            self.close()

    def _build_rates_map(self, payload: Dict[str, object]) -> rates_importer.RatesMap:
        rows = payload.get("rows") or []
        result: rates_importer.RatesMap = {}
        for row in rows:
            if not isinstance(row, dict):
                continue
            src = row.get("source")
            tgt = row.get("target")
            if not src or not tgt:
                continue

            src_code = rates_importer._normalize_language(str(src))
            tgt_code = rates_importer._normalize_language(str(tgt))

            basic = self._coerce_float(row.get("basic"))
            complex_ = self._coerce_float(row.get("complex"))
            hour = self._coerce_float(row.get("hour"))

            result[(src_code, tgt_code)] = {
                "basic": 0.0 if basic is None else float(basic),
                "complex": 0.0 if complex_ is None else float(complex_),
                "hour": 0.0 if hour is None else float(hour),
            }

        return result

    @staticmethod
    def _coerce_float(value) -> Optional[float]:
        if value is None:
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            return None
