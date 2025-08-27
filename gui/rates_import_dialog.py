"""Dialog for importing rate tables from Excel files."""

from __future__ import annotations

from typing import Iterable, List, Tuple

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
        main_layout.addLayout(file_layout)

        options_layout = QGridLayout()
        options_layout.addWidget(QLabel("Currency:"), 0, 0)
        self.currency_combo = QComboBox(self)
        self.currency_combo.addItems(["USD", "EUR", "RUB", "CNY"])
        options_layout.addWidget(self.currency_combo, 0, 1)

        options_layout.addWidget(QLabel("Rates:"), 1, 0)
        self.rate_combo = QComboBox(self)
        self.rate_combo.addItems(["R1", "R2"])
        options_layout.addWidget(self.rate_combo, 1, 1)

        self.load_btn = QPushButton("Load")
        self.load_btn.clicked.connect(self._load)
        options_layout.addWidget(self.load_btn, 2, 0, 1, 2)

        main_layout.addLayout(options_layout)

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

    def _browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel", "", "Excel Files (*.xlsx)")
        if path:
            self.file_edit.setText(path)

    def _load(self) -> None:
        path = self.file_edit.text()
        if not path:
            return
        currency = self.currency_combo.currentText()
        rate_type = self.rate_combo.currentText()
        try:
            rates = rates_importer.load_rates_from_excel(path, currency, rate_type)
        except Exception as exc:  # pragma: no cover - GUI feedback
            self.table.setRowCount(0)
            self.table.setRowCount(1)
            self.table.setSpan(0, 0, 1, self.table.columnCount())
            item = QTableWidgetItem(str(exc))
            self.table.setItem(0, 0, item)
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

        self.table.setRowCount(len(matches))
        for row, match in enumerate(matches):
            self.table.setItem(row, 0, QTableWidgetItem(match.gui_source))
            self.table.setItem(row, 1, QTableWidgetItem(match.gui_target))
            self.table.setItem(row, 2, QTableWidgetItem(match.excel_source))
            self.table.setItem(row, 3, QTableWidgetItem(match.excel_target))
            if match.rates:
                self.table.setItem(row, 4, QTableWidgetItem(str(match.rates["basic"])))
                self.table.setItem(row, 5, QTableWidgetItem(str(match.rates["complex"])))
                self.table.setItem(row, 6, QTableWidgetItem(str(match.rates["hour"])))
            else:
                for col in range(4, 7):
                    self.table.setItem(row, col, QTableWidgetItem(""))

    def selected_rates(self) -> List[rates_importer.PairMatch]:
        """Return the current mapping as edited by the user."""
        matches: List[rates_importer.PairMatch] = []
        for row in range(self.table.rowCount()):
            gui_src = self.table.item(row, 0).text()
            gui_tgt = self.table.item(row, 1).text()
            excel_src = self.table.item(row, 2).text()
            excel_tgt = self.table.item(row, 3).text()
            basic = self.table.item(row, 4).text()
            complex_ = self.table.item(row, 5).text()
            hour = self.table.item(row, 6).text()
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
