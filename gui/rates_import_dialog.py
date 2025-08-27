"""Dialog for importing rate tables from Excel files."""

from __future__ import annotations

from typing import Dict, Iterable, List, Tuple

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from logic import rates_importer


class ExcelRatesDialog(QDialog):
    """Dialog allowing the user to load and review rate tables."""

    def __init__(self, gui_pairs: Iterable[Tuple[str, str]], parent=None) -> None:
        super().__init__(parent)
        self._pairs = list(gui_pairs)
        self.setWindowTitle("Import Rates")

        main_layout = QVBoxLayout(self)

        file_layout = QHBoxLayout()
        self.file_edit = QLineEdit(self)
        browse_btn = QPushButton("...")
        browse_btn.clicked.connect(self._browse)
        file_layout.addWidget(QLabel("Excel file:"))
        file_layout.addWidget(self.file_edit)
        file_layout.addWidget(browse_btn)
        self.file_edit.editingFinished.connect(self._load)
        main_layout.addLayout(file_layout)

        options_layout = QGridLayout()
        options_layout.addWidget(QLabel("Currency:"), 0, 0)
        self.currency_combo = QComboBox(self)
        self.currency_combo.addItems(["USD", "EUR", "RUB", "CNY"])
        self.currency_combo.currentTextChanged.connect(self._load)
        options_layout.addWidget(self.currency_combo, 0, 1)

        options_layout.addWidget(QLabel("Rates:"), 1, 0)
        self.rate_combo = QComboBox(self)
        self.rate_combo.addItems(["R1", "R2"])
        self.rate_combo.currentTextChanged.connect(self._load)
        options_layout.addWidget(self.rate_combo, 1, 1)

        options_layout.addWidget(QLabel("Use rate:"), 2, 0)
        self.apply_combo = QComboBox(self)
        self.apply_combo.addItems(["Basic", "Complex"])
        options_layout.addWidget(self.apply_combo, 2, 1)

        main_layout.addLayout(options_layout)

        self.status_label = QLabel("")
        main_layout.addWidget(self.status_label)

        self.table = QTableWidget(0, 7, self)
        headers = [
            "GUI Source",
            "GUI Target",
            "Excel Source",
            "Excel Target",
            "Basic",
            "Complex",
            "Hour",
        ]
        self.table.setHorizontalHeaderLabels(headers)
        main_layout.addWidget(self.table)

        self.resize(800, 400)
        self._rates: Dict[Tuple[str, str], Dict[str, float]] = {}
        self._name_to_code: Dict[str, str] = {}

    def showEvent(self, event):  # pragma: no cover - GUI behaviour
        super().showEvent(event)
        if self.file_edit.text():
            self._load()

    def _browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel", "", "Excel Files (*.xlsx)")
        if path:
            self.file_edit.setText(path)
            self._load()

    def _load(self) -> None:
        path = self.file_edit.text()
        if not path:
            self.status_label.setText("No file selected")
            self.table.setRowCount(0)
            return
        currency = self.currency_combo.currentText()
        rate_type = self.rate_combo.currentText()
        try:
            rates = rates_importer.load_rates_from_excel(path, currency, rate_type)
            self._rates = rates
        except Exception as exc:  # pragma: no cover - GUI feedback
            self.table.setRowCount(0)
            self.status_label.setText(str(exc))
            return

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
            self.table.setItem(row, 0, QTableWidgetItem(match.gui_source))
            self.table.setItem(row, 1, QTableWidgetItem(match.gui_target))

            src_combo = QComboBox(self.table)
            src_combo.addItems(lang_names)
            src_combo.setCurrentText(match.excel_source)
            src_combo.currentTextChanged.connect(
                lambda _t, r=row: self._update_rate_from_row(r)
            )
            self.table.setCellWidget(row, 2, src_combo)

            tgt_combo = QComboBox(self.table)
            tgt_combo.addItems(lang_names)
            tgt_combo.setCurrentText(match.excel_target)
            tgt_combo.currentTextChanged.connect(
                lambda _t, r=row: self._update_rate_from_row(r)
            )
            self.table.setCellWidget(row, 3, tgt_combo)

            if match.rates:
                self.table.setItem(row, 4, QTableWidgetItem(str(match.rates["basic"])))
                self.table.setItem(row, 5, QTableWidgetItem(str(match.rates["complex"])))
                self.table.setItem(row, 6, QTableWidgetItem(str(match.rates["hour"])))
            else:
                for col in range(4, 7):
                    self.table.setItem(row, col, QTableWidgetItem(""))

        self.status_label.setText(f"Loaded {len(rates)} rate pairs")

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
        for col in range(4, 7):
            item = self.table.item(row, col)
            if item is None:
                item = QTableWidgetItem()
                self.table.setItem(row, col, item)
        if rate:
            self.table.item(row, 4).setText(str(rate["basic"]))
            self.table.item(row, 5).setText(str(rate["complex"]))
            self.table.item(row, 6).setText(str(rate["hour"]))
        else:
            for col in range(4, 7):
                self.table.item(row, col).setText("")

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
            if basic or complex_ or hour:
                try:
                    rate = {
                        "basic": float(basic or 0),
                        "complex": float(complex_ or 0),
                        "hour": float(hour or 0),
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
