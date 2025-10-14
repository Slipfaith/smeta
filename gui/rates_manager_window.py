"""Embedded version of the standalone rates utility."""

from __future__ import annotations

from functools import partial
from typing import Dict, Iterable, List, Optional, Tuple

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from gui.styles import (
    RATES_IMPORT_DIALOG_STYLE,
    STATUS_LABEL_DEFAULT_STYLE,
    STATUS_LABEL_SUCCESS_STYLE,
)
from logic import rates_importer
from logic.translation_config import tr
from rates1 import RateTab

RateRow = Dict[str, Optional[float]]


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

        self._setup_ui()
        self.setStyleSheet(RATES_IMPORT_DIALOG_STYLE)

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        controls_layout = QHBoxLayout()
        controls_layout.setSpacing(8)

        self.auto_update_checkbox = QCheckBox(
            tr("Автоматически подставлять ставки", self._lang())
        )
        self.auto_update_checkbox.setChecked(True)
        controls_layout.addWidget(self.auto_update_checkbox)

        self.auto_fill_btn = QPushButton(tr("Подставить ставки сейчас", self._lang()))
        self.auto_fill_btn.clicked.connect(lambda: self.auto_fill_from_rates(force=True))
        controls_layout.addWidget(self.auto_fill_btn)

        controls_layout.addStretch()
        layout.addLayout(controls_layout)

        apply_layout = QHBoxLayout()
        apply_layout.setSpacing(8)

        apply_label = QLabel(tr("Применить", self._lang()) + ":")
        apply_layout.addWidget(apply_label)

        self.apply_combo = QComboBox()
        self.apply_combo.addItems(["Basic", "Complex"])
        self.apply_combo.setFixedWidth(110)
        apply_layout.addWidget(self.apply_combo)
        apply_layout.addStretch()

        layout.addLayout(apply_layout)

        status_layout = QHBoxLayout()
        status_layout.addStretch()
        self.status_label = QLabel(tr("Данные не загружены", self._lang()))
        self.status_label.setStyleSheet(STATUS_LABEL_DEFAULT_STYLE)
        status_layout.addWidget(self.status_label)
        layout.addLayout(status_layout)

        self.table = QTableWidget(0, 7)
        headers = [
            "Source",
            "Target",
            "Excel Source",
            "Excel Target",
            "Basic",
            "Complex",
            "Hour",
        ]
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        header.setSectionResizeMode(5, QHeaderView.Fixed)
        header.setSectionResizeMode(6, QHeaderView.Fixed)
        self.table.setColumnWidth(0, 140)
        self.table.setColumnWidth(1, 140)
        self.table.setColumnWidth(4, 80)
        self.table.setColumnWidth(5, 80)
        self.table.setColumnWidth(6, 80)
        layout.addWidget(self.table, 1)

        import_layout = QHBoxLayout()
        import_layout.addStretch()
        self.import_btn = QPushButton(tr("Импортировать в программу", self._lang()))
        self.import_btn.clicked.connect(self.import_requested.emit)
        import_layout.addWidget(self.import_btn)
        layout.addLayout(import_layout)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def set_pairs(self, pairs: Iterable[Tuple[str, str]]) -> None:
        self._pairs = list(pairs)
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
        self._apply_matches(auto_fill=self.auto_update_checkbox.isChecked())
        self._update_status()

    def auto_fill_from_rates(self, force: bool = False) -> None:
        allow_force = force or self.auto_update_checkbox.isChecked()
        for row in range(self.table.rowCount()):
            self._update_rate_from_row(row, force=allow_force)

    def selected_rates(self) -> List[rates_importer.PairMatch]:
        matches: List[rates_importer.PairMatch] = []
        for row in range(self.table.rowCount()):
            gui_src = self._safe_item_text(row, 0)
            gui_tgt = self._safe_item_text(row, 1)

            excel_src = self._combo_text(row, 2)
            excel_tgt = self._combo_text(row, 3)

            basic = self._safe_item_text(row, 4)
            complex_ = self._safe_item_text(row, 5)
            hour = self._safe_item_text(row, 6)

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
        return self.apply_combo.currentText().lower()

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

        for row, (src, tgt) in enumerate(self._pairs):
            src_item = self.table.item(row, 0)
            if src_item is None:
                src_item = QTableWidgetItem(src)
                src_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.table.setItem(row, 0, src_item)
            else:
                src_item.setText(src)

            tgt_item = self.table.item(row, 1)
            if tgt_item is None:
                tgt_item = QTableWidgetItem(tgt)
                tgt_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.table.setItem(row, 1, tgt_item)
            else:
                tgt_item.setText(tgt)

            for column in (4, 5, 6):
                item = self.table.item(row, column)
                if item is None:
                    item = QTableWidgetItem("")
                    item.setTextAlignment(Qt.AlignCenter)
                    self.table.setItem(row, column, item)

            self._ensure_combo(row, 2)
            self._ensure_combo(row, 3)

    def _refresh_language_combos(self) -> None:
        for row in range(self.table.rowCount()):
            for column in (2, 3):
                combo = self._ensure_combo(row, column)
                previous = combo.currentText()
                combo.blockSignals(True)
                combo.clear()
                combo.addItems(self._lang_names)
                if previous:
                    self._set_combo_text(combo, previous)
                combo.blockSignals(False)

    def _apply_matches(self, auto_fill: bool) -> None:
        if not self._pairs:
            return

        matches = rates_importer.match_pairs(self._pairs, self._rates)
        for row, match in enumerate(matches):
            src_combo = self._ensure_combo(row, 2)
            tgt_combo = self._ensure_combo(row, 3)
            self._set_combo_text(src_combo, match.excel_source)
            self._set_combo_text(tgt_combo, match.excel_target)

            if auto_fill and match.rates:
                self._apply_rate_values(row, match.rates)

    def _set_combo_text(self, combo: QComboBox, text: str) -> None:
        if not text:
            return
        index = combo.findText(text)
        if index == -1:
            combo.addItem(text)
            index = combo.findText(text)
        combo.setCurrentIndex(index if index >= 0 else 0)

    def _ensure_combo(self, row: int, column: int) -> QComboBox:
        widget = self.table.cellWidget(row, column)
        if isinstance(widget, QComboBox):
            return widget
        combo = QComboBox()
        combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        combo.currentTextChanged.connect(partial(self._handle_combo_change, row))
        self.table.setCellWidget(row, column, combo)
        return combo

    def _handle_combo_change(self, row: int, _text: str) -> None:
        self._update_rate_from_row(row)

    def _update_rate_from_row(self, row: int, force: bool = False) -> None:
        if not force and not self.auto_update_checkbox.isChecked():
            return

        src_combo = self.table.cellWidget(row, 2)
        tgt_combo = self.table.cellWidget(row, 3)
        if not isinstance(src_combo, QComboBox) or not isinstance(tgt_combo, QComboBox):
            return

        src_code = self._name_to_code.get(src_combo.currentText())
        tgt_code = self._name_to_code.get(tgt_combo.currentText())
        if not src_code or not tgt_code:
            return

        rate = self._rates.get((src_code, tgt_code))
        if rate:
            self._apply_rate_values(row, rate)
        elif force or self.auto_update_checkbox.isChecked():
            self._clear_rate_values(row)

    def _apply_rate_values(self, row: int, rate: RateRow) -> None:
        basic = rate.get("basic")
        complex_ = rate.get("complex")
        hour = rate.get("hour")
        self._set_rate_text(row, 4, basic)
        self._set_rate_text(row, 5, complex_)
        self._set_rate_text(row, 6, hour)

    def _clear_rate_values(self, row: int) -> None:
        for column in (4, 5, 6):
            item = self.table.item(row, column)
            if item is not None:
                item.setText("")

    def _set_rate_text(self, row: int, column: int, value: Optional[float]) -> None:
        item = self.table.item(row, column)
        if item is None:
            item = QTableWidgetItem("")
            item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, column, item)
        if value is None:
            item.setText("")
        else:
            item.setText(f"{value:.2f}")

    def _safe_item_text(self, row: int, column: int) -> str:
        item = self.table.item(row, column)
        return item.text() if item else ""

    def _combo_text(self, row: int, column: int) -> str:
        widget = self.table.cellWidget(row, column)
        return widget.currentText() if isinstance(widget, QComboBox) else ""

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


class RatesManagerWindow(QMainWindow):
    """Main window embedding the remote rates explorer."""

    def __init__(self, main_window) -> None:
        super().__init__(main_window)
        self._main_window = main_window
        self._current_pairs: List[Tuple[str, str]] = []

        self.setWindowTitle(tr("Панель ставок", main_window.gui_lang))
        self.resize(1400, 720)

        central = QWidget()
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(10, 10, 10, 10)
        root_layout.setSpacing(12)

        self.rate_tab = RateTab()
        root_layout.addWidget(self.rate_tab, 1)

        self.mapping_widget = RatesMappingWidget(lambda: self._main_window.gui_lang, self)
        root_layout.addWidget(self.mapping_widget, 1)

        self.setCentralWidget(central)

        self.rate_tab.rates_updated.connect(self._handle_rate_payload)
        self.mapping_widget.import_requested.connect(self._apply_to_main_window)

    # ------------------------------------------------------------------
    # Public helpers
    # ------------------------------------------------------------------
    def update_pairs(self, pairs: Iterable[Tuple[str, str]]) -> None:
        self._current_pairs = list(pairs)
        self.mapping_widget.set_pairs(self._current_pairs)

    # ------------------------------------------------------------------
    # Slots
    # ------------------------------------------------------------------
    def _handle_rate_payload(self, payload: Dict[str, object]) -> None:
        rates = self._build_rates_map(payload)
        source_label = str(payload.get("source_label", ""))
        currency = str(payload.get("currency", ""))
        rate_type = str(payload.get("rate_type", ""))
        self.mapping_widget.set_rates_data(rates, source_label, currency, rate_type)

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
        self._main_window._apply_rates_from_matches(matches, rate_key)

    # ------------------------------------------------------------------
    # Data helpers
    # ------------------------------------------------------------------
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
