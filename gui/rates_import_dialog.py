"""Dialog for importing rate tables from Excel files."""

from __future__ import annotations

from typing import Dict, Iterable, List, Optional, Tuple

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from logic import online_rates, rates_importer
from gui.styles import (
    RATES_IMPORT_DIALOG_APPLY_COMBO_WIDTH,
    RATES_IMPORT_DIALOG_BROWSE_BUTTON_WIDTH,
    RATES_IMPORT_DIALOG_BUTTON_LAYOUT_MARGINS,
    RATES_IMPORT_DIALOG_CURRENCY_COMBO_WIDTH,
    RATES_IMPORT_DIALOG_MAIN_MARGINS,
    RATES_IMPORT_DIALOG_MAIN_SPACING,
    RATES_IMPORT_DIALOG_RATE_COMBO_WIDTH,
    RATES_IMPORT_DIALOG_SECTION_SPACING,
    RATES_IMPORT_DIALOG_SIZE,
    RATES_IMPORT_DIALOG_STYLE,
    RATES_IMPORT_DIALOG_TABLE_COLUMN_WIDTHS,
    STATUS_LABEL_DEFAULT_STYLE,
    STATUS_LABEL_SUCCESS_STYLE,
    STATUS_LABEL_ERROR_STYLE,
)


class ExcelRatesDialog(QDialog):
    """Dialog allowing the user to load and review rate tables."""

    apply_requested = Signal()

    def __init__(self, gui_pairs: Iterable[Tuple[str, str]], parent=None) -> None:
        super().__init__(parent)
        self._pairs = list(gui_pairs)
        self.setWindowTitle("Import Rates")
        self.setFixedSize(*RATES_IMPORT_DIALOG_SIZE)

        self._setup_ui()
        self._setup_styles()

        self._rates: Dict[Tuple[str, str], Dict[str, float]] = {}
        self._name_to_code: Dict[str, str] = {}
        self._remote_sources = online_rates.available_sources()
        self._remote_cache: Dict[Tuple[str, str, str], rates_importer.RatesMap] = {}
        self._current_source: Optional[str] = None

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setSpacing(RATES_IMPORT_DIALOG_MAIN_SPACING)
        layout.setContentsMargins(*RATES_IMPORT_DIALOG_MAIN_MARGINS)

        source_layout = QHBoxLayout()
        source_layout.setSpacing(RATES_IMPORT_DIALOG_SECTION_SPACING)
        source_layout.addWidget(QLabel("Source:"))
        self.source_combo = QComboBox()
        self.source_combo.addItem("Local File", userData=None)
        for key, remote in self._remote_sources.items():
            self.source_combo.addItem(remote.label, userData=key)
        self.source_combo.currentIndexChanged.connect(self._on_source_changed)
        source_layout.addWidget(self.source_combo, 1)
        source_layout.addStretch()
        layout.addLayout(source_layout)

        file_layout = QHBoxLayout()
        file_layout.setSpacing(RATES_IMPORT_DIALOG_SECTION_SPACING)

        file_layout.addWidget(QLabel("File:"))
        self.file_edit = QLineEdit()
        self.file_edit.setPlaceholderText("Select Excel file...")
        self.file_edit.editingFinished.connect(self._load)
        file_layout.addWidget(self.file_edit, 1)

        self.browse_btn = QPushButton("Browse")
        self.browse_btn.setFixedWidth(RATES_IMPORT_DIALOG_BROWSE_BUTTON_WIDTH)
        self.browse_btn.clicked.connect(self._browse)
        file_layout.addWidget(self.browse_btn)

        layout.addLayout(file_layout)
        self._on_source_changed(self.source_combo.currentIndex())

        config_layout = QGridLayout()
        config_layout.setSpacing(RATES_IMPORT_DIALOG_SECTION_SPACING)

        config_layout.addWidget(QLabel("Currency:"), 0, 0)
        self.currency_combo = QComboBox()
        self.currency_combo.addItems(["USD", "EUR", "RUB", "CNY"])
        self.currency_combo.setFixedWidth(RATES_IMPORT_DIALOG_CURRENCY_COMBO_WIDTH)
        self.currency_combo.currentTextChanged.connect(self._load)
        config_layout.addWidget(self.currency_combo, 0, 1)

        config_layout.addWidget(QLabel("Type:"), 0, 2)
        self.rate_combo = QComboBox()
        self.rate_combo.addItems(["R1", "R2"])
        self.rate_combo.setFixedWidth(RATES_IMPORT_DIALOG_RATE_COMBO_WIDTH)
        self.rate_combo.currentTextChanged.connect(self._load)
        config_layout.addWidget(self.rate_combo, 0, 3)

        config_layout.addWidget(QLabel("Apply:"), 0, 4)
        self.apply_combo = QComboBox()
        self.apply_combo.addItems(["Basic", "Complex"])
        self.apply_combo.setFixedWidth(RATES_IMPORT_DIALOG_APPLY_COMBO_WIDTH)
        config_layout.addWidget(self.apply_combo, 0, 5)

        config_layout.setColumnStretch(6, 1)

        self.status_label = QLabel("No file selected")
        self.status_label.setAlignment(Qt.AlignRight)
        config_layout.addWidget(self.status_label, 0, 7)

        layout.addLayout(config_layout)

        self.table = QTableWidget(0, 7)
        headers = ["Source", "Target", "Excel Source", "Excel Target", "Basic", "Complex", "Hour"]
        self.table.setHorizontalHeaderLabels(headers)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        header.setSectionResizeMode(5, QHeaderView.Fixed)
        header.setSectionResizeMode(6, QHeaderView.Fixed)

        for column, width in RATES_IMPORT_DIALOG_TABLE_COLUMN_WIDTHS.items():
            self.table.setColumnWidth(column, width)

        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)

        layout.addWidget(self.table, 1)

        btn_layout = QHBoxLayout()
        btn_layout.setContentsMargins(*RATES_IMPORT_DIALOG_BUTTON_LAYOUT_MARGINS)
        btn_layout.addStretch()
        apply_btn = QPushButton("Apply")
        apply_btn.clicked.connect(self.apply_requested.emit)
        btn_layout.addWidget(apply_btn)
        layout.addLayout(btn_layout)

    def _setup_styles(self) -> None:
        self.setStyleSheet(RATES_IMPORT_DIALOG_STYLE)

    def _on_source_changed(self, _index: int) -> None:
        source_key = self.source_combo.currentData()
        self._current_source = source_key
        is_local = source_key is None
        self.file_edit.setEnabled(is_local)
        self.browse_btn.setEnabled(is_local)
        if is_local:
            self.file_edit.setPlaceholderText("Select Excel file...")
        else:
            self.file_edit.setText(self.source_combo.currentText())
        self._load()

    def showEvent(self, event):
        super().showEvent(event)
        if self.file_edit.text():
            self._load()

    def _browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self.file_edit.setText(path)
            self._load()

    def _load(self) -> None:
        source_key = getattr(self, 'source_combo', None)
        source_key = source_key.currentData() if source_key else None
        currency = self.currency_combo.currentText()
        rate_type = self.rate_combo.currentText()

        rates: Optional[rates_importer.RatesMap] = None
        status_message: str

        if source_key is None:
            path = self.file_edit.text().strip()
            if not path:
                self._rates = {}
                self.table.setRowCount(0)
                self.status_label.setText("No file")
                self.status_label.setStyleSheet(STATUS_LABEL_DEFAULT_STYLE)
                return
            try:
                rates = rates_importer.load_rates_from_excel(path, currency, rate_type)
            except Exception as exc:
                self._rates = {}
                self.table.setRowCount(0)
                self.status_label.setText(f"Error: {str(exc)[:40]}...")
                self.status_label.setStyleSheet(STATUS_LABEL_ERROR_STYLE)
                return
            status_message = f"Loaded {len(rates)} pairs"
        else:
            cache_key = (source_key, currency, rate_type)
            try:
                rates = self._remote_cache.get(cache_key)
                if rates is None:
                    rates = online_rates.load_remote_rates(source_key, currency, rate_type)
                    self._remote_cache[cache_key] = rates
            except Exception as exc:
                self._rates = {}
                self.table.setRowCount(0)
                self.status_label.setText(f"Error: {str(exc)[:40]}...")
                self.status_label.setStyleSheet(STATUS_LABEL_ERROR_STYLE)
                return
            source_label = self.source_combo.currentText()
            status_message = f"Loaded {len(rates)} pairs from {source_label}"

        self._rates = rates or {}
        self.status_label.setText(status_message)
        self.status_label.setStyleSheet(STATUS_LABEL_SUCCESS_STYLE)

        self.table.clearSpans()

        if self._pairs:
            matches: List[rates_importer.PairMatch] = rates_importer.match_pairs(self._pairs, rates)
        else:
            matches = [
                rates_importer.PairMatch(
                    gui_source=rates_importer._language_name(src),
                    gui_target=rates_importer._language_name(tgt),
                    excel_source=rates_importer._language_name(src),
                    excel_target=rates_importer._language_name(tgt),
                    rates=rate,
                )
                for (src, tgt), rate in rates.items()
            ]

        lang_codes = set()
        for src_code, tgt_code in rates:
            lang_codes.add(src_code)
            lang_codes.add(tgt_code)
        self._name_to_code = {
            rates_importer._language_name(code): code for code in lang_codes
        }
        lang_names = sorted(self._name_to_code.keys())

        self.table.setRowCount(len(matches))
        for row, match in enumerate(matches):
            gui_src_item = QTableWidgetItem(match.gui_source)
            gui_src_item.setFlags(gui_src_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 0, gui_src_item)

            gui_tgt_item = QTableWidgetItem(match.gui_target)
            gui_tgt_item.setFlags(gui_tgt_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 1, gui_tgt_item)

            src_combo = QComboBox()
            src_combo.addItems(lang_names)
            src_combo.setCurrentText(match.excel_source)
            src_combo.currentTextChanged.connect(
                lambda _t, r=row: self._update_rate_from_row(r)
            )
            self.table.setCellWidget(row, 2, src_combo)

            tgt_combo = QComboBox()
            tgt_combo.addItems(lang_names)
            tgt_combo.setCurrentText(match.excel_target)
            tgt_combo.currentTextChanged.connect(
                lambda _t, r=row: self._update_rate_from_row(r)
            )
            self.table.setCellWidget(row, 3, tgt_combo)

            if match.rates:
                basic_item = QTableWidgetItem(f"{match.rates['basic']:.2f}")
                complex_item = QTableWidgetItem(f"{match.rates['complex']:.2f}")
                hour_item = QTableWidgetItem(f"{match.rates['hour']:.2f}")
            else:
                basic_item = QTableWidgetItem("")
                complex_item = QTableWidgetItem("")
                hour_item = QTableWidgetItem("")

            basic_item.setFlags(basic_item.flags() & ~Qt.ItemIsEditable)
            complex_item.setFlags(complex_item.flags() & ~Qt.ItemIsEditable)
            hour_item.setFlags(hour_item.flags() & ~Qt.ItemIsEditable)

            self.table.setItem(row, 4, basic_item)
            self.table.setItem(row, 5, complex_item)
            self.table.setItem(row, 6, hour_item)

    def _update_rate_from_row(self, row: int) -> None:
        if not self._rates:
            return
        src_w = self.table.cellWidget(row, 2)
        tgt_w = self.table.cellWidget(row, 3)
        if not isinstance(src_w, QComboBox) or not isinstance(tgt_w, QComboBox):
            return
        src_code = rates_importer._normalize_language(src_w.currentText())
        tgt_code = rates_importer._normalize_language(tgt_w.currentText())
        rate = self._rates.get((src_code, tgt_code))

        if rate:
            self.table.item(row, 4).setText(f"{rate['basic']:.2f}")
            self.table.item(row, 5).setText(f"{rate['complex']:.2f}")
            self.table.item(row, 6).setText(f"{rate['hour']:.2f}")
        else:
            self.table.item(row, 4).setText("")
            self.table.item(row, 5).setText("")
            self.table.item(row, 6).setText("")

    def selected_rates(self) -> List[rates_importer.PairMatch]:
        """Return the current mapping as edited by the user."""
        matches: List[rates_importer.PairMatch] = []
        for row in range(self.table.rowCount()):
            gui_src = self.table.item(row, 0).text()
            gui_tgt = self.table.item(row, 1).text()
            src_w = self.table.cellWidget(row, 2)
            tgt_w = self.table.cellWidget(row, 3)
            excel_src = (
                src_w.currentText() if isinstance(src_w, QComboBox) else ""
            )
            excel_tgt = (
                tgt_w.currentText() if isinstance(tgt_w, QComboBox) else ""
            )
            basic_item = self.table.item(row, 4)
            complex_item = self.table.item(row, 5)
            hour_item = self.table.item(row, 6)
            basic = basic_item.text() if basic_item else ""
            complex_ = complex_item.text() if complex_item else ""
            hour = hour_item.text() if hour_item else ""
            rate = None
            if basic and complex_ and hour:
                try:
                    rate = {
                        "basic": float(basic),
                        "complex": float(complex_),
                        "hour": float(hour),
                    }
                except ValueError:
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